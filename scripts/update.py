#!/usr/bin/env python3
"""Descarga, conserva originales y genera un HTML autónomo con doce indicadores."""
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
from statistics import mean, median
from urllib.request import Request, urlopen
from urllib.parse import urljoin, urlparse, urlencode
from zoneinfo import ZoneInfo

import openpyxl
import pdfplumber
import xlrd

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
        elif sid=='cost_m2':
            if 'Costo de construcción del metro cuadrado (en pesos)' not in title:
                raise ValueError('Cambió el encabezado del costo por m²')
            for sheet in w:
                if not re.fullmatch(r'20\d{2}',sheet.title): continue
                data=list(sheet.values)
                if data[1][1]!='Total Ciudad' or tuple(data[2][2:6])!=('Tipo I','Tipo II','Tipo III','Tipo IV'):
                    raise ValueError('Cambió la estructura del costo por m²: '+sheet.title)
                for row in data[3:]:
                    m=month(row[0])
                    if m:
                        put(out,f'{sheet.title}-{m:02d}',row[1:6],'*' in str(row[0]))
                        if any(v<=0 for v in row[1:6]): raise ValueError('Costo por m² no positivo')
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
        elif sid=='deeds':
            expected=('Cantidad de actos de compraventa','Monto de escrituras de compraventa','Cantidad de actos hipotecarios','Monto escrituras hipotecarias')
            if 'compraventa' not in title or tuple(str(x).strip() for x in rows[1][1:5])!=expected:
                raise ValueError('Cambió la estructura de escrituras')
            year=None
            for row in rows[2:]:
                if isinstance(row[0],int): year=row[0]; continue
                m=month(row[0])
                if m and year and any(v is not None for v in row[1:5]):
                    put(out,f'{year}-{m:02d}',row[1:5],'*' in str(row[0]))
                    if any(v<0 for v in row[1:5]): raise ValueError('Escrituras negativas')
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

def es_number(text):
    """Número con formato argentino: punto de miles y coma decimal."""
    return float(text.replace('.','').replace(',','.'))

def parse_wp_posts(path):
    """Informes mensuales del Colegio de Escribanos publicados como entradas de WordPress."""
    posts=json.loads(path.read_text())
    if not isinstance(posts,list): raise ValueError('Respuesta inesperada del Colegio de Escribanos')
    out={}; chosen={}
    for p in sorted(posts,key=lambda p:p['date']):
        title=html.unescape(p['title']['rendered'])
        match=re.search(r'Cantidad de escrituras de compraventa realizadas en\s+([A-Za-zé]+)\s+(?:de\s+)?(\d{4})',title,re.I)
        if not match or not month(match[1]): continue
        key=f'{match[2]}-{month(match[1]):02d}'
        text=html.unescape(re.sub(r'<[^>]+>',' ',p['content']['rendered']))
        text=re.sub(r'\s+',' ',re.sub(r'\[/?et_pb[^\]]*\]',' ',text))
        count=re.search(r'(?:Actos de escrituras|Escrituras) de compraventa:?\s*([\d.]+)',text)
        if not count: continue
        amount=re.search(r'Monto involucrado:?\s*\$?\s*([\d.]+(?:,\d+)?)\s*millones',text)
        # La cifra de hipotecas del informe corresponde al mes del informe (su acumulado anual lo confirma),
        # aunque a veces el texto nombre el mes anterior.
        mortgages=re.search(r'([\d.]+)\s+escrituras formalizadas con hipoteca',text) or re.search(r'con hipoteca(?: bancaria)? totalizaron\s+([\d.]+)',text)
        value=es_number(count[1])
        if not 0<=value<50000: raise ValueError('Cantidad de escrituras fuera de rango: '+key)
        # Si un mes tiene más de un informe, vale el publicado último.
        chosen[key]=p['date']
        out[key]={'values':[value],'amount':es_number(amount[1])*1e6 if amount else None,
                  'mortgages':es_number(mortgages[1]) if mortgages else None,'provisional':False,'published':p['date'][:10]}
    if len(out)<100: raise ValueError('El Colegio de Escribanos no contiene el histórico esperado')
    return dict(sorted(out.items())),'Escrituras de compraventa · informes mensuales del Colegio de Escribanos'

def parse_bcra(path,aggregate,as_of):
    raw=json.loads(path.read_text())
    detail=raw.get('detalle') if isinstance(raw,dict) else None
    if not detail: raise ValueError('Respuesta inesperada del BCRA')
    buckets=defaultdict(list); seen=set()
    for r in sorted(detail,key=lambda r:r['fecha']):
        day=date.fromisoformat(r['fecha'])
        if day>as_of: continue
        if r['fecha'] in seen: raise ValueError('Dato duplicado del BCRA: '+r['fecha'])
        seen.add(r['fecha'])
        value=numeric(r.get('valor'))
        if value is None or value<=0: raise ValueError('Dato inválido del BCRA: '+r['fecha'])
        buckets[r['fecha'][:7]].append(value)
    pick=mean if aggregate=='mean' else (lambda v:v[-1])
    out={k:{'values':[pick(v)],'provisional':k==as_of.strftime('%Y-%m'),'observations':len(v)} for k,v in sorted(buckets.items())}
    return out,raw.get('descripcion','Serie del BCRA')

def parse_cbt(path):
    book=xlrd.open_workbook(str(path))
    sheet=book.sheet_by_index(0)
    title=str(sheet.cell_value(0,0))
    if 'adulto equivalente' not in title or 'Canasta básica total' not in str(sheet.row_values(3)):
        raise ValueError('Cambió la estructura de la canasta básica')
    out={}
    for i in range(sheet.nrows):
        row=sheet.row_values(i)
        if isinstance(row[0],float) and row[0]>30000:
            day=xlrd.xldate_as_datetime(row[0],book.datemode)
            put(out,day.strftime('%Y-%m'),[row[1],row[3]])
    if len(out)<100: raise ValueError('La canasta básica no contiene el histórico esperado')
    return out,'Canasta básica alimentaria y total por adulto equivalente · INDEC (GBA)'

def parse_csv_rate(path):
    rows=list(csv.reader(path.read_text().splitlines()))
    if rows[0][0]!='observation_date': raise ValueError('Cambió el formato de FRED')
    out={}
    for d,v in rows[1:]:
        if v in ('','.'): continue
        put(out,d[:7],[float(v)])
    if len(out)<100: raise ValueError('La tasa de EE.UU. no contiene el histórico esperado')
    return out,'Rendimiento del bono del Tesoro de EE.UU. a 10 años · promedio mensual'

def parse_source(s,path,settings,as_of):
    fmt=s['format']
    if fmt=='xlsx': return parse_xlsx(s,path)
    if fmt=='pdf': return parse_ripte(path)
    if fmt=='wp_posts': return parse_wp_posts(path)
    if fmt=='bcra':
        rows,_=parse_bcra(path,s['aggregate'],as_of)
        return rows,s.get('title',s['id'])
    if fmt=='xls': return parse_cbt(path)
    if fmt=='csv': return parse_csv_rate(path)
    return parse_fx(path,settings['exchange_rate'],as_of)

class Links(HTMLParser):
    def __init__(self): super().__init__(); self.links=[]
    def handle_starttag(self,tag,attrs):
        if tag=='a':
            url=dict(attrs).get('href','')
            if url: self.links.append(url)

def fetch(url):
    request=Request(url,headers={'User-Agent':'MonitorInmobiliarioLocal/1.0','Accept':'*/*'})
    with urlopen(request,timeout=40) as r:
        content=r.read(20_000_001)
        if len(content)>20_000_000: raise ValueError('Descarga demasiado grande')
        return content

def fetch_json(url):
    request=Request(url,headers={'User-Agent':'MonitorInmobiliarioLocal/1.0','Accept':'application/json'})
    with urlopen(request,timeout=40) as r:
        content=r.read(20_000_001)
        if len(content)>20_000_000: raise ValueError('Descarga demasiado grande')
        return json.loads(content),r.headers

def download_wp_posts(s):
    posts=[]; page=1
    while True:
        query=urlencode({'search':s['search'],'per_page':100,'page':page,'_fields':'id,date,modified,link,title,content'})
        batch,headers=fetch_json(s['url']+'?'+query)
        if not isinstance(batch,list): raise ValueError('Respuesta inesperada del Colegio de Escribanos')
        posts+=[p for p in batch if 'Cantidad de escrituras de compraventa' in html.unescape(p['title']['rendered'])]
        if page>=int(headers.get('X-WP-TotalPages','1')) or page>=20: break
        page+=1
    posts.sort(key=lambda p:p['id'])
    return json.dumps(posts,ensure_ascii=False,sort_keys=True).encode()

def download_bcra(s,as_of):
    detail=[]; offset=0
    while True:
        query=urlencode({'desde':s['since'],'hasta':as_of.isoformat(),'limit':3000,'offset':offset})
        data,_=fetch_json(s['url']+'?'+query)
        results=data.get('results') or []
        if len(results)!=1: raise ValueError('Respuesta inesperada del BCRA')
        detail+=results[0]['detalle']
        total=data['metadata']['resultset']['count']
        offset+=3000
        if offset>=total or offset>30000: break
    detail.sort(key=lambda r:r['fecha'])
    return json.dumps({'idVariable':results[0]['idVariable'],'detalle':detail},sort_keys=True).encode()

def refresh_one(s,old,settings,as_of,now):
    try:
        url=s['url']
        if s['format']=='wp_posts': content=download_wp_posts(s)
        elif s['format']=='bcra': content=download_bcra(s,as_of)
        elif s['format'] in ('json','xls','csv'): content=fetch(url)
        else:
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

def fmt_pct(value):
    return f'{value:g}'.replace('.',',')+'%'

def shift_month(key,n):
    y,m=map(int,key.split('-')); total=y*12+m-1+n
    return f'{total//12}-{total%12+1:02d}'

def usable_fx(fx,key):
    return key in fx and fx[key].get('observations',0)>=10

def rolling(points,window=12):
    """Promedio móvil de `window` meses consecutivos; no se calcula si falta un mes."""
    by={k:(v,p) for k,v,p in points}; out=[]
    for k in by:
        months=[shift_month(k,-i) for i in range(window)]
        if all(m in by for m in months):
            out.append([k,mean(by[m][0] for m in months),any(by[m][1] for m in months)])
    return out

def merge_deeds(idecba,colegio,tolerance):
    """Escrituras por mes. El Colegio de Escribanos es la fuente primaria; IDECBA, que publica los
    mismos datos, completa los meses sin informe. Se descartan los años de IDECBA cuya cantidad
    difiere del Colegio más que `tolerance` (mediana de los meses comunes)."""
    ratios=defaultdict(list)
    for k,r in colegio.items():
        if k in idecba and r['values'][0]>0: ratios[k[:4]].append(idecba[k]['values'][0]/r['values'][0])
    rejected=sorted(y for y,v in ratios.items() if abs(median(v)-1)>tolerance)
    merged={}
    for k in sorted(set(idecba)|set(colegio)):
        ide=idecba.get(k) if k[:4] not in rejected else None
        ce=colegio.get(k)
        if ce:
            merged[k]={'sales':ce['values'][0],'amount':ce['amount'] if ce['amount'] is not None else (ide['values'][1] if ide else None),
                       'mortgages':ce['mortgages'] if ce['mortgages'] is not None else (ide['values'][2] if ide else None),
                       'source':'Colegio','provisional':ce['provisional']}
        elif ide:
            merged[k]={'sales':ide['values'][0],'amount':ide['values'][1],'mortgages':ide['values'][2],'source':'IDECBA','provisional':ide['provisional']}
    return merged,rejected

def french_payment(principal,annual_rate_pct,years):
    r=annual_rate_pct/100/12; n=years*12
    return principal*r/(1-(1+r)**-n) if r else principal/n

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
    costs=data['cost_m2']
    eligible=[k for k in costs if k in fx and fx[k]['observations']>=10]
    labels=['Total Ciudad','Tipo I · sencilla','Tipo II · sencilla con local','Tipo III · confortable','Tipo IV · suntuosa']
    usd=[line(n,[[k,costs[k]['values'][i]/fx[k]['values'][0],costs[k]['provisional'] or fx[k]['provisional']] for k in eligible]) for i,n in enumerate(labels)]
    changes=[]
    for k,value,provisional in usd[0]['points']:
        y,m=map(int,k.split('-')); prev=f'{y-(m==1)}-{(m-2)%12+1:02d}'
        if prev in eligible:
            previous=costs[prev]['values'][0]/fx[prev]['values'][0]
            changes.append([k,100*(value/previous-1),provisional or costs[prev]['provisional'] or fx[prev]['provisional']])
    c.append({'id':8,'title':'Costo de construcción por m² en USD MEP','intro':'Costo mensual · Total Ciudad y cuatro modelos de vivienda · CABA','sources':['cost_m2','fx'],'note':'Costo de construcción en pesos / dólar MEP vendedor promedio del mismo mes. Se usa el Total Ciudad oficial, no el promedio simple de los cuatro modelos. Incluye materiales, mano de obra y gastos generales; excluye terreno, IVA, derechos y permisos, costos financieros y beneficio de la constructora. Son edificios de vivienda colectiva, no un presupuesto para una casa particular ni el precio de venta. Sólo meses cerrados con al menos diez cotizaciones; no se interpolan faltantes. La historia en USD empieza con la cobertura MEP disponible. * indica datos provisorios.','views':[chart('Costo por m² · USD MEP','USD MEP / m²',usd[:1],'Mensual'),chart('Por tipo de vivienda · USD MEP','USD MEP / m²',usd,'Mensual'),chart('Variación mensual · USD MEP','%',[line('Total Ciudad',changes)],'Mensual'),chart('Costo por m² · pesos','ARS / m²',[line('Total Ciudad',[point(costs,k) for k in costs])],'Mensual')]})
    fxb=data['fx_blue']; blue_label=settings['long_exchange_label']
    areas=settings['rent_areas_m2']; rooms=['1 ambiente','2 ambientes','3 ambientes']
    # 7 · Rentabilidad por tipología, con referencias de tasa.
    ust=data['ust10']; rate=settings['mortgage_rate_pct']
    for kind,prefix in [('usados','used'),('a estrenar','new')]:
        lines=[]
        for i in range(3):
            pts=[]
            for k,r in data[f'{prefix}{i+1}'].items():
                months=quarter_months(k)
                if all(m in rent and usable_fx(fx,m) for m in months):
                    monthly=mean(rent[m]['values'][i]/fx[m]['values'][0] for m in months)
                    pts.append([k,100*12*monthly/(r['values'][0]*areas[i]),r['provisional'] or any(rent[m]['provisional'] for m in months)])
            lines.append(line(rooms[i],pts))
        span=sorted({p[0] for ln in lines for p in ln['points']})
        lines.append(line('Bono EE.UU. 10 años',[[k,mean(ust[m]['values'][0] for m in quarter_months(k)),False] for k in span if all(m in ust for m in quarter_months(k))]))
        lines.append(line(f'Tasa UVA supuesta ({fmt_pct(rate)})',[[k,rate,False] for k in span]))
        c[6]['views'].append(chart(f'Por tipología · {kind}','% anual',lines,'Trimestral'))
    c[6]['intro']='Estimación bruta · 1, 2 y 3 ambientes · dólar MEP'
    c[6]['sources']+=['new2','new3','used2','used3','ust10']
    c[6]['note']+=(f' Las vistas por tipología usan el alquiler de 1, 2 y 3 ambientes sobre unidades base de {areas[0]}, {areas[1]} y {areas[2]} m² y el precio publicado del mismo tipo. '
                   f'Referencias: rendimiento nominal en USD del bono del Tesoro de EE.UU. a 10 años (FRED) y la tasa real del crédito UVA supuesta en el gráfico 10 ({fmt_pct(rate)}); no son comparables uno a uno.')
    # 8 · Costo de construcción con dólar blue: historia desde 2015.
    c[7]['views'].append(chart('Costo por m² · USD blue','USD blue / m²',[line('Total Ciudad',[[k,costs[k]['values'][0]/fxb[k]['values'][0],costs[k]['provisional'] or fxb[k]['provisional']] for k in costs if usable_fx(fxb,k)])],'Mensual'))
    c[7]['note']+=f' La vista en dólar blue ({blue_label.lower()}, mercado informal) permite ver la historia desde 2015, antes de la cobertura MEP.'
    # 9 · Escrituras de compraventa.
    deeds,rejected=merge_deeds(data['deeds'],data['deeds_ce'],settings['deeds_tolerance'])
    sales=[[k,r['sales'],r['provisional']] for k,r in deeds.items()]
    def avg_deed(rate_series):
        return [[k,r['amount']/r['sales']/rate_series[k]['values'][0],r['provisional'] or rate_series[k]['provisional']] for k,r in deeds.items() if r['amount'] and r['sales']>=100 and usable_fx(rate_series,k)]
    rejected_text=(' Se descartaron los datos de IDECBA de '+', '.join(rejected)+' porque difieren del Colegio en más de '+fmt_pct(100*settings['deeds_tolerance'])+'.') if rejected else ''
    c.append({'id':9,'title':'Escrituras de compraventa','intro':'Actos por mes · todos los inmuebles · CABA','sources':['deeds_ce','deeds','fx'],
              'note':'Cantidad de escrituras de compraventa de inmuebles ubicados en la Ciudad, de todo tipo (no sólo departamentos). Fuente primaria: informes mensuales del Colegio de Escribanos; los meses sin informe se completan con la serie de IDECBA, elaborada con los mismos datos.'+rejected_text+
                      ' La escritura se firma en general uno o dos meses después de la reserva. Monto promedio = monto declarado / cantidad de escrituras / dólar promedio del mes: no es precio por m² ni precio de cierre de departamentos. Se omiten meses con menos de 100 escrituras (abril de 2020).',
              'views':[chart('Compraventas por mes','Escrituras',[line('Compraventas',sales),line('Promedio móvil 12 meses',rolling(sales))],'Mensual'),
                       chart('Monto promedio · USD MEP','USD por escritura',[line('Monto promedio',avg_deed(fx))],'Mensual'),
                       chart('Monto promedio · USD blue','USD por escritura',[line('Monto promedio',avg_deed(fxb))],'Mensual')]})
    # 10 · Cuota de un crédito UVA frente al alquiler y al salario.
    uva=data['uva']; ub=settings['uva_base']; ltv=settings['mortgage_ltv']; yrs=settings['mortgage_years']
    if not (ub in uva and ub in rent and ub in wage): raise ValueError('Falta el mes base de UVA')
    idx=lambda series,col=0:[[k,100*series[k]['values'][col]/series[ub]['values'][col],series[k]['provisional']] for k in series if k>=ub and k in uva]
    def relative(a,b):
        other={k:(v,p) for k,v,p in b}
        return [[k,100*v/other[k][0],p or other[k][1]] for k,v,p in a if k in other]
    payment_rent=[]; payment_wage=[]
    for name,sid in [('2 ambientes usado','used2'),('2 ambientes a estrenar','new2')]:
        pr=[]; pw=[]
        for k,r in data[sid].items():
            months=quarter_months(k)
            if not all(usable_fx(fx,m) and m in rent and m in wage for m in months): continue
            principal=r['values'][0]*areas[1]*mean(fx[m]['values'][0] for m in months)*ltv
            payment=french_payment(principal,rate,yrs)
            provisional=r['provisional'] or any(rent[m]['provisional'] for m in months)
            pr.append([k,100*payment/mean(rent[m]['values'][1] for m in months),provisional])
            pw.append([k,100*payment/mean(wage[m]['values'][0] for m in months),provisional])
        payment_rent.append(line(name,pr)); payment_wage.append(line(name,pw))
    c.append({'id':10,'title':'Cuota UVA frente al alquiler','intro':f'2 ambientes · crédito de {fmt_pct(100*ltv)} del valor a {yrs} años · UVA + {fmt_pct(rate)}','sources':['uva','rent','ripte','used2','new2','fx'],
              'note':f'Base 100 = {ub}, primer mes completo de la UVA. La cuota de un crédito UVA por sistema francés es fija en UVA, así que en pesos sigue al valor de la UVA (BCRA). El alquiler es el publicado para 2 ambientes (unidad base de {areas[1]} m², IDECBA) y el salario es RIPTE. '
                      'Índice relativo = índice de la cuota / índice del alquiler (o del salario) × 100: 100 significa que se movieron igual desde la base; más de 100, que la cuota subió más. '
                      f'Las otras vistas calculan la cuota inicial de un crédito nuevo en cada trimestre: precio publicado del m² × {areas[1]} m² × dólar MEP × {fmt_pct(100*ltv)}, a {yrs} años y tasa real de {fmt_pct(rate)} constante (supuesto; las tasas reales cambiaron y hubo años casi sin crédito). Excluye seguros, gastos e impuestos.',
              'views':[chart(f'Cuota frente a alquiler y salario · {ub} = 100','Índice relativo',[line('Cuota UVA / alquiler 2 ambientes',relative(idx(uva),idx(rent,1))),line('Cuota UVA / salario RIPTE',relative(idx(uva),idx(wage)))],'Mensual'),
                       chart(f'Cuota, alquiler y salario · base {ub} = 100','Índice',[line('Cuota de un crédito UVA',idx(uva)),line('Alquiler 2 ambientes',idx(rent,1)),line('Salario RIPTE',idx(wage))],'Mensual'),
                       chart('Cuota inicial como % del alquiler','%',payment_rent,'Trimestral'),
                       chart('Cuota inicial como % del salario RIPTE','%',payment_wage,'Trimestral')]})
    # 11 · Salario y canasta básica en dólares.
    cbt=data['cbt']
    def in_usd(series,rates,col):
        return [[k,series[k]['values'][col]/rates[k]['values'][0],series[k]['provisional'] or rates[k]['provisional']] for k in series if usable_fx(rates,k)]
    c.append({'id':11,'title':'Salario y canasta básica en dólares','intro':'RIPTE y canasta básica total por adulto equivalente · promedio mensual del dólar','sources':['ripte','cbt','fx'],
              'note':'Indicador de si la Argentina está "cara" o "barata" en dólares. RIPTE: remuneración imponible promedio de trabajadores estables (nacional). Canasta básica total: línea de pobreza por adulto equivalente del INDEC (GBA), no una canasta de la Ciudad. Cada valor mensual en pesos se divide por el promedio del dólar del mismo mes. '
                      f'La vista en dólar blue ({blue_label.lower()}, mercado informal) cubre el salario desde 2011 y la canasta desde abril de 2016 (primer dato publicado), antes de la cobertura MEP. Canastas por salario = RIPTE / canasta básica total, sin tipo de cambio.',
              'views':[chart('En dólar MEP','USD por mes',[line('Salario RIPTE',in_usd(wage,fx,0)),line('Canasta básica total',in_usd(cbt,fx,1))],'Mensual'),
                       chart('En dólar blue','USD por mes',[line('Salario RIPTE',in_usd(wage,fxb,0)),line('Canasta básica total',in_usd(cbt,fxb,1))],'Mensual'),
                       chart('Canastas por salario','Canastas',[line('RIPTE / canasta básica total',[[k,wage[k]['values'][0]/cbt[k]['values'][1],wage[k]['provisional'] or cbt[k]['provisional']] for k in cbt if k in wage])],'Mensual')]})
    # 12 · Crédito hipotecario.
    mortgage=data['mortgage']
    with_mortgage=[[k,r['mortgages'],r['provisional']] for k,r in deeds.items() if r['mortgages'] is not None]
    share=[[k,100*r['mortgages']/r['sales'],r['provisional']] for k,r in deeds.items() if r['mortgages'] is not None and r['sales']>=100]
    real_base=max(k for k in mortgage if k in ipc)
    c.append({'id':12,'title':'Crédito hipotecario','intro':'Escrituras con hipoteca en CABA y saldo de préstamos hipotecarios en el país','sources':['deeds_ce','deeds','mortgage','ipc','fx'],
              'note':'Escrituras con hipoteca: informes del Colegio de Escribanos (escrituras de compraventa formalizadas con hipoteca); los meses sin ese dato se completan con los actos hipotecarios de IDECBA, que pueden diferir algunos puntos. '
                      f'Saldo: préstamos hipotecarios al sector privado de todo el país (BCRA), último dato de cada mes, en dólar MEP promedio del mes o deflactado por IPCBA a pesos de {real_base}. El saldo cae cuando se amortiza más de lo que se presta. No incluye mora.',
              'views':[chart('Escrituras con hipoteca','Escrituras',[line('Con hipoteca',with_mortgage),line('Promedio móvil 12 meses',rolling(with_mortgage))],'Mensual'),
                       chart('% de compraventas con hipoteca','%',[line('Con hipoteca',share)],'Mensual'),
                       chart('Saldo de préstamos hipotecarios · USD MEP','USD millones',[line('Saldo',[[k,mortgage[k]['values'][0]/fx[k]['values'][0],mortgage[k]['provisional'] or fx[k]['provisional']] for k in mortgage if usable_fx(fx,k)])],'Mensual'),
                       chart(f'Saldo real · ARS de {real_base}','ARS millones a precios constantes',[line('Saldo',[[k,mortgage[k]['values'][0]*ipc[real_base]['values'][0]/ipc[k]['values'][0],mortgage[k]['provisional']] for k in mortgage if k in ipc and k<=real_base])],'Mensual')]})
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
    fx_path=ROOT/manifest['fx']['path']
    blue,_=parse_fx(fx_path,settings['long_exchange_rate'],as_of)
    ds['fx_blue']=({k:v for k,v in blue.items() if k<as_of.strftime('%Y-%m')},'Cotización vendedora · '+settings['long_exchange_rate'])
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
    print(f'Generados {len(cards)} gráficos: '+str(folder/'index.html'))
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
