#!/usr/bin/env python3
"""Descarga, conserva originales y genera un HTML autónomo con siete indicadores."""
import argparse
import csv
import hashlib
import html
import json
import math
import re
import sys
import warnings
from collections import defaultdict
from contextlib import contextmanager
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, date
from html.parser import HTMLParser
from pathlib import Path
from statistics import mean
from urllib.request import Request, urlopen
from urllib.parse import urljoin, urlparse
from zoneinfo import ZoneInfo

import openpyxl
import pdfplumber

ROOT = Path(__file__).resolve().parents[1]
MONTHS = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre']
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def month(value):
    text=str(value or '').strip().lower()
    for i,name in enumerate(MONTHS,1):
        if text.startswith(name) or text[:3] == name[:3]:
            return i
    return None

def numeric(value):
    if isinstance(value,(int,float)) and not isinstance(value,bool) and math.isfinite(value):
        return float(value)
    return None

def put(out,key,values,provisional=False):
    if key in out:
        raise ValueError('Período duplicado: '+key)
    if not values or any(numeric(v) is None for v in values):
        raise ValueError('Dato faltante o no numérico: '+key)
    out[key]={'values':[float(v) for v in values],'provisional':provisional}

@contextmanager
def workbook(path):
    w=openpyxl.load_workbook(path,data_only=True,read_only=True)
    try: yield w
    finally: w.close()

def parse_xlsx(source,path):
    out={}
    with workbook(path) as w:
        rows=list(w.worksheets[0].values)
        title=str(rows[0][0])
        sid=source['id']
        if sid=='icc':
            if 'Costo de la Construcción' not in title: raise ValueError('Cambió el encabezado de ICC')
            for sheet in w:
                if not re.fullmatch(r'20\d{2}',sheet.title): continue
                data=list(sheet.values)
                if tuple(str(x).strip() for x in data[2][1:5]) != ('Nivel general','Materiales','Mano de obra','Gastos generales'):
                    raise ValueError('Cambió la estructura de ICC: '+sheet.title)
                for row in data[3:]:
                    m=month(row[0])
                    if m: put(out,f'{sheet.title}-{m:02d}',row[1:5],'*' in str(row[0]))
        elif sid=='rent':
            if 'alquiler' not in title or [str(x) for x in rows[2][2:5]] != ['1 ambiente','2 ambientes','3 ambientes']:
                raise ValueError('Cambió la estructura de alquileres')
            year=None
            for row in rows[3:]:
                if isinstance(row[0],int): year=row[0]
                m=month(row[1])
                if m and year: put(out,f'{year}-{m:02d}',row[2:5],'*' in str(row[1]))
        elif sid=='days':
            if 'Tiempo medio' not in title or rows[1][2:6] != ('Total','1 ambiente','2 ambientes','3 ambientes'):
                raise ValueError('Cambió la estructura de tiempo de publicación')
            year=None
            for row in rows[2:]:
                if isinstance(row[0],int): year=row[0]
                q=re.match(r'([1-4]).*trim',str(row[1] or ''))
                if q and year: put(out,f'{year}-{int(q[1])*3:02d}',row[2:6],'*' in str(row[1]))
        elif sid=='ipc':
            if 'empalmado' not in title: raise ValueError('Se esperaba IPCBA empalmado')
            for row in rows:
                if isinstance(row[0],(date,datetime)): put(out,row[0].strftime('%Y-%m'),[row[1]],'*' in str(row[0]))
        else:
            is_area=sid.startswith('area')
            if (is_area and 'Superficie total' not in title) or (not is_area and 'metro cuadrado' not in title):
                raise ValueError('Cambió el encabezado: '+sid)
            total=next((r for r in rows if str(r[0]).strip().lower()=='total'),None)
            if total is None: raise ValueError('No hay Total Ciudad')
            year=None
            for col in range(1,len(rows[1])):
                if isinstance(rows[1][col],int): year=rows[1][col]
                period=rows[2][col]
                if is_area: m=month(period)
                else:
                    q=re.match(r'([1-4]).*trim',str(period or ''))
                    m=int(q[1])*3 if q else None
                if m and year: put(out,f'{year}-{m:02d}',[total[col]],'*' in str(period))
    if not out: raise ValueError('Serie vacía: '+source['id'])
    return dict(sorted(out.items())),title

def parse_ripte(path):
    out={}
    year=None
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for line in (page.extract_text() or '').splitlines():
                if re.fullmatch(r'(19|20)\d{2}',line.strip()): year=int(line.strip())
                match=re.match(r'^([A-Za-z]+)\s+\$\s*([\d.,]+)',line.strip())
                if match and year:
                    m=month(match[1])
                    if m: put(out,f'{year}-{m:02d}',[float(match[2].replace('.','').replace(',','.'))])
    if not out or len(out)<100: raise ValueError('RIPTE no contiene el histórico esperado')
    keys=sorted(out)
    for a,b in zip(keys,keys[1:]):
        y,m=map(int,a.split('-'))
        expected=f'{y+(m==12)}-{m%12+1:02d}'
        if b != expected: raise ValueError('Hueco en RIPTE entre '+a+' y '+b)
    return dict(sorted(out.items())),'RIPTE · remuneración imponible promedio de trabajadores estables'

def parse_fx(path,house,as_of):
    raw=json.loads(path.read_text())
    if not isinstance(raw,list): raise ValueError('Respuesta inesperada de cotizaciones')
    buckets=defaultdict(list)
    seen=set()
    for r in raw:
        if r.get('casa')!=house: continue
        day=date.fromisoformat(r['fecha'])
        if day>as_of or day.weekday()>4: continue
        if r['fecha'] in seen: raise ValueError('Cotización duplicada: '+r['fecha'])
        seen.add(r['fecha'])
        value=numeric(r.get('venta'))
        if value is None or value<=0: raise ValueError('Cotización inválida')
        buckets[r['fecha'][:7]].append(value)
    out={k:{'values':[mean(v)],'provisional':k==as_of.strftime('%Y-%m'),'observations':len(v)} for k,v in sorted(buckets.items())}
    if not out: raise ValueError('No hay cotizaciones para '+house)
    return out,'Cotización vendedora · '+house+' · promedio de días de semana publicados'

def parse_source(s,path,settings,as_of):
    if s['format']=='xlsx': return parse_xlsx(s,path)
    if s['format']=='pdf': return parse_ripte(path)
    return parse_fx(path,settings['exchange_rate'],as_of)

class Links(HTMLParser):
    def __init__(self): super().__init__(); self.links=[]
    def handle_starttag(self,tag,attrs):
        if tag=='a':
            url=dict(attrs).get('href','')
            if url: self.links.append(url)

def fetch(url):
    request=Request(url,headers={'User-Agent':'MonitorInmobiliarioLocal/1.0','Accept':'*/*'})
    with urlopen(request,timeout=20) as r:
        content=r.read(20_000_001)
        if len(content)>20_000_000: raise ValueError('Descarga demasiado grande')
        return content

def refresh_one(s,old,settings,as_of,now):
    try:
        url=s['url']
        if s['format']!='json':
            links=Links(); links.feed(fetch(s['page']).decode('utf-8'))
            if s['format']=='xlsx': candidates=[u for u in links.links if re.search(r'\.xlsx(?:\?|$)',u,re.I)]
            else: candidates=[u for u in links.links if 'ripte' in u.lower() and re.search(r'\.pdf(?:\?|$)',u,re.I)]
            if len(set(candidates))!=1: raise ValueError('No se identificó un único archivo de descarga')
            link=candidates[0]
            # RIPTE publica ocasionalmente un href "blank:#/sites/default/files/...".
            # Usar sólo la ruta de archivo bajo el mismo sitio oficial, sin interpretar el esquema.
            if s['id']=='ripte' and link.startswith('blank:#/sites/default/files/'):
                link=link[len('blank:#'):]
            url=urljoin(s['page'],link)
            if (urlparse(url).hostname or '').removeprefix('www.') != (urlparse(s['page']).hostname or '').removeprefix('www.'):
                raise ValueError('La descarga cambió de dominio; revisar fuente')
        content=fetch(url)
        digest=hashlib.sha256(content).hexdigest()
        relative=Path('data/raw')/s['id']/(digest[:16]+Path(s['filename']).suffix)
        candidate=ROOT/'data'/'incoming'/(s['id']+Path(s['filename']).suffix)
        candidate.parent.mkdir(parents=True,exist_ok=True); candidate.write_bytes(content)
        parsed,_=parse_source(s,candidate,settings,as_of)
        if old:
            previous,_=parse_source(s,ROOT/old['path'],settings,as_of)
            if max(parsed)<max(previous) or len(parsed)<len(previous):
                raise ValueError('La descarga pierde historia o retrocede en fecha')
        dest=ROOT/relative; dest.parent.mkdir(parents=True,exist_ok=True)
        if not dest.exists(): dest.write_bytes(content)
        return {**s,'url':url,'path':str(relative),'sha256':digest,'downloaded_at':old['downloaded_at'] if old and digest==old['sha256'] else now,'checked_at':now,'status':'sin cambios' if old and digest==old['sha256'] else 'actualizado'}
    except Exception as exc:
        if not old: raise
        return {**old,'checked_at':now,'status':'no se pudo actualizar','error':str(exc)}

def point(series,key,col=0,factor=1):
    r=series[key]
    return [key,r['values'][col]*factor,r['provisional']]

def line(name,points,**extra): return {'name':name,'points':points,**extra}

def chart(title,unit,lines,frequency,averages=False):
    return {'title':title,'unit':unit,'lines':lines,'frequency':frequency,'averages':averages}

def quarter_months(key):
    y,m=map(int,key.split('-'))
    return [f'{y}-{n:02d}' for n in range(m-2,m+1)]

def build_charts(ds,settings):
    data={s:rows for s,(rows,_) in ds.items()}
    c=[]
    names=['Nivel general','Materiales','Mano de obra','Gastos generales']
    icc=data['icc']; ipc=data['ipc']; base=settings['index_base']
    normalized=[line(n,[point(icc,k,i,100/icc[base]['values'][i]) for k in icc if k>=base]) for i,n in enumerate(names)]
    normalized.append(line('IPCBA',[point(ipc,k,0,100/ipc[base]['values'][0]) for k in ipc if base<=k<=max(icc)]))
    variations=[]
    for i,n in enumerate(names):
        pts=[]
        for k in icc:
            y,m=map(int,k.split('-')); prev=f'{y-(m==1)}-{(m-2)%12+1:02d}'
            if prev in icc: pts.append([k,100*(icc[k]['values'][i]/icc[prev]['values'][i]-1),icc[k]['provisional']])
        variations.append(line(n,pts))
    c.append({'id':1,'title':'Costos de construcción','intro':'CABA · materiales, mano de obra y gastos generales','sources':['icc','ipc'],'note':'Índices normalizados a enero de 2021 = 100 para compararlos con IPCBA. La otra vista muestra variaciones mensuales calculadas a partir de los índices. * indica datos provisorios.','views':[chart('Evolución frente a la inflación','Índice · ene 2021 = 100',normalized,'Mensual'),chart('Variación mensual','%',variations,'Mensual')]})
    days=data['days']
    c.append({'id':2,'title':'Tiempo medio de publicación','intro':'Departamentos en venta · CABA','sources':['days'],'note':'Días transcurridos entre la primera publicación y la fecha de análisis del conjunto de avisos. No mide el tiempo hasta concretar una venta. Promedio histórico: media de los valores trimestrales disponibles por serie.','views':[chart('Permanencia de avisos','Días',[line(n,[point(days,k,i) for k in days]) for i,n in enumerate(['Total','1 ambiente','2 ambientes','3 ambientes'])],'Trimestral',True)]})
    area=[line(f'{i} ambiente'+('s' if i>1 else ''),[point(data[f'area{i}'],k) for k in data[f'area{i}']]) for i in range(1,4)]
    c.append({'id':3,'title':'Superficie publicada en alquiler','intro':'Usados y a estrenar · 1, 2 y 3 ambientes · CABA','sources':['area1','area2','area3'],'note':'Superficie cubierta total de los avisos de alquiler. Se usa el Total Ciudad oficial, sin sumar barrios ni convertir m² en cantidad de departamentos.','views':[chart('Oferta de alquileres','m²',area,'Mensual')]})
    rent=data['rent']; base_real=max(set(rent)&set(ipc)); real=[]; nominal=[]
    for i in range(3):
        name=f'{i+1} ambiente'+('s' if i else '')
        real.append(line(name,[[k,rent[k]['values'][i]*ipc[base_real]['values'][0]/ipc[k]['values'][0],rent[k]['provisional']] for k in rent if k in ipc and k<=base_real]))
        nominal.append(line(name,[point(rent,k,i) for k in rent]))
    c.append({'id':4,'title':'Alquileres a precios constantes','intro':f'Pesos de {base_real} · unidades base de 30, 43 y 70 m²','sources':['rent','ipc'],'note':f'Alquiler real = alquiler publicado × IPCBA de {base_real} / IPCBA del mes. La fuente usa superficies estándar de 30, 43 y 70 m² para 1, 2 y 3 ambientes; no son el alquiler medio de cualquier departamento.','views':[chart('Alquiler real','ARS a precios constantes',real,'Mensual'),chart('Alquiler nominal','ARS',nominal,'Mensual')]})
    price_start=settings['price_history_start']
    prices=[line(f'{i} ambiente'+('s' if i>1 else ''),[point(data[f'new{i}'],k) for k in data[f'new{i}'] if k>=price_start]) for i in range(1,4)]
    c.append({'id':5,'title':'Precio del m² a estrenar','intro':'Precio publicado en USD · 1, 2 y 3 ambientes · CABA','sources':['new1','new2','new3'],'note':'Total Ciudad: precio de oferta, no precio de cierre. Frecuencia trimestral de los archivos públicos. Promedios históricos calculados desde 2017 sobre los trimestres disponibles; no se inventan observaciones mensuales.','views':[chart('Precio de venta','USD / m²',prices,'Trimestral',True)]})
    fx=data['fx']; wage=data['ripte']; ratios=[]; yields=[]; years=[]
    for sid,name in [('new1','A estrenar'),('used1','Usado')]:
        rp=[]; yp=[]; ap=[]
        for k,r in data[sid].items():
            months=quarter_months(k)
            if all(m in wage and m in fx and fx[m]['observations']>=10 for m in months):
                salary=mean(wage[m]['values'][0]/fx[m]['values'][0] for m in months)
                rp.append([k,r['values'][0]/salary,r['provisional']])
            if all(m in rent and m in fx and fx[m]['observations']>=10 for m in months):
                monthly=mean(rent[m]['values'][0]/fx[m]['values'][0] for m in months)
                result=12*monthly/(r['values'][0]*settings['rent_area_m2'])
                provisional=r['provisional'] or any(rent[m]['provisional'] for m in months)
                yp.append([k,result*100,provisional]); ap.append([k,1/result,provisional])
        ratios.append(line(name,rp)); yields.append(line(name,yp)); years.append(line(name,ap))
    c.append({'id':6,'title':'Salarios para comprar 1 m²','intro':'Monoambientes · RIPTE convertido al dólar MEP','sources':['new1','used1','ripte','fx'],'note':'Precio del m² / promedio trimestral de RIPTE mensual en USD MEP. Sólo trimestres con los tres meses de RIPTE y de cotizaciones. RIPTE es remuneración imponible de trabajadores estables, no el salario de toda la economía. Se muestran usados y a estrenar porque el video no precisa el tipo.','views':[chart('Acceso al m²','Salarios RIPTE',ratios,'Trimestral')]})
    c.append({'id':7,'title':'Rentabilidad y recuperación','intro':'Estimación bruta · monoambiente de 30 m² · dólar MEP','sources':['new1','used1','rent','fx'],'note':'Estimación comparable, no réplica exacta del video: el alquiler publicado mezcla usados y a estrenar. Se aplica el mismo alquiler de una unidad base de 30 m² a ambos precios de compra. Rendimiento = 12 × promedio trimestral del alquiler mensual en USD / (precio del m² × 30). Años = 100 / rendimiento en %. Excluye vacancia, impuestos, mantenimiento, gastos de compra y apreciación.','views':[chart('Rendimiento bruto anualizado','% anual',yields,'Trimestral'),chart('Años para recuperar la inversión','Años',years,'Trimestral')]})
    for card in c:
        if any(not ln['points'] for view in card['views'] for ln in view['lines']):
            raise ValueError('Sin datos para el gráfico '+str(card['id']))
    return c

def atomic_json(path,value):
    path.parent.mkdir(parents=True,exist_ok=True)
    tmp=path.with_suffix(path.suffix+'.tmp')
    tmp.write_text(json.dumps(value,ensure_ascii=False,indent=2,allow_nan=False)+'\n')
    tmp.replace(path)

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--refresh',action='store_true',help='Consultar fuentes antes de generar')
    parser.add_argument('--source',help='Actualizar sólo una fuente por ID')
    parser.add_argument('--as-of',type=date.fromisoformat,help='Fecha de corte YYYY-MM-DD')
    args=parser.parse_args()
    settings=json.loads((ROOT/'config/settings.json').read_text())
    now=datetime.now(ZoneInfo(settings['timezone'])); as_of=args.as_of or now.date()
    now_text=now.isoformat(timespec='seconds')
    sources=json.loads((ROOT/'config/sources.json').read_text())
    manifest=json.loads((ROOT/'data/manifest.json').read_text())
    if args.source and args.source not in {s['id'] for s in sources}: parser.error('Fuente inexistente')
    if args.refresh:
        selected=[s for s in sources if not args.source or s['id']==args.source]
        with ThreadPoolExecutor(max_workers=6) as pool:
            futures={pool.submit(refresh_one,s,manifest.get(s['id']),settings,as_of,now_text):s for s in selected}
            for future in as_completed(futures):
                s=futures[future]; manifest[s['id']]=future.result()
                print(s['id']+': '+manifest[s['id']]['status'],flush=True)
        atomic_json(ROOT/'data/manifest.json',manifest)
    ds={}
    for s in sources:
        meta=manifest[s['id']]; path=ROOT/meta['path']
        if hashlib.sha256(path.read_bytes()).hexdigest()!=meta['sha256']: raise ValueError('El original fue alterado: '+s['id'])
        rows,title=parse_source(s,path,settings,as_of)
        # Exclude incomplete/current monthly periods and not-yet-ended quarters.
        cutoff=as_of.strftime('%Y-%m')
        rows={k:v for k,v in rows.items() if k<cutoff}
        if not rows: raise ValueError('No hay historia cerrada: '+s['id'])
        ds[s['id']]=(rows,title)
        meta.update({'title':title,'first_period':min(rows),'last_period':max(rows),'observations':len(rows)})
    cards=build_charts(ds,settings)
    failures=[sid for sid,m in manifest.items() if m.get('status')=='no se pudo actualizar']
    report={'generated_at':now_text,'as_of':as_of.isoformat(),'refresh_requested':args.refresh,'settings':settings,'sources':manifest,'cards':cards,'refresh_failures':failures}
    atomic_json(ROOT/'data/processed/series.json',{sid:{'title':title,'records':rows} for sid,(rows,title) in ds.items()})
    atomic_json(ROOT/'data/processed/report.json',report)
    folder=ROOT/'output'; folder.mkdir(exist_ok=True)
    with (folder/'series.csv').open('w',newline='',encoding='utf-8-sig') as f:
        writer=csv.writer(f); writer.writerow(['grafico','vista','serie','periodo','valor','unidad','provisorio'])
        for card in cards:
            for view in card['views']:
                for ln in view['lines']:
                    for k,v,p in ln['points']: writer.writerow([card['id'],view['title'],ln['name'],k,v,view['unit'],p])
    template=(ROOT/'templates/report.html').read_text()
    payload=json.dumps(report,ensure_ascii=False,allow_nan=False).replace('<','\\u003c')
    rendered=template.replace('REPORT_PAYLOAD',payload)
    (folder/'index.html.tmp').write_text(rendered)
    (folder/'index.html.tmp').replace(folder/'index.html')
    run_id=now.strftime('%Y%m%d-%H%M%S-%f')
    atomic_json(ROOT/'data/runs'/(run_id+'.json'),{'generated_at':now_text,'as_of':as_of.isoformat(),'refresh_requested':args.refresh,'failures':failures,'source_hashes':{k:v['sha256'] for k,v in manifest.items()}})
    print('Generados 7 gráficos: '+str(folder/'index.html'))
    print('Períodos: '+', '.join(sid+' '+max(rows) for sid,(rows,_) in ds.items()))
    if failures:
        print('AVISO: se conservaron originales anteriores para: '+', '.join(failures),file=sys.stderr)
        return 2
    return 0

if __name__=='__main__':
    try: sys.exit(main())
    except Exception as exc:
        print('No se reemplazó el HTML: '+str(exc),file=sys.stderr)
        sys.exit(1)
