"""Comprobaciones de cobertura, cifras de fuente y cálculos temporales."""
import copy
import importlib.util
import json
import unittest
from datetime import date
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
# Planillas guardadas en la primera publicación de cada fuente, contra las que se contrastaron
# a mano las cifras de control. Las fuentes revisan meses recientes en publicaciones
# posteriores, así que los controles leen estas copias fijas y no la última descarga.
CONTROL_SNAPSHOTS={
    'icc':'data/raw/icc/52c930aeaedb33fa.xlsx',
    'rent':'data/raw/rent/ac3e2de8e8c2c5f4.xlsx',
    'ripte':'data/raw/ripte/3e87f594ae7a215d.pdf',
    'new1':'data/raw/new1/34017f66b3ba8f40.xlsx',
    'cost_m2':'data/raw/cost_m2/f708835148156f0d.xlsx',
    'deeds_ce':'data/raw/deeds_ce/cef9f128f9a974ca.json',
    'cbt':'data/raw/cbt/c040c3e37a8574ad.xls',
    'uva':'data/raw/uva/4c08773b13abac85.json',
}
spec=importlib.util.spec_from_file_location('monitor',ROOT/'scripts/update.py')
m=importlib.util.module_from_spec(spec); spec.loader.exec_module(m)

class PipelineTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.settings=json.loads((ROOT/'config/settings.json').read_text())
        sources=json.loads((ROOT/'config/sources.json').read_text())
        manifest=json.loads((ROOT/'data/manifest.json').read_text())
        cls.ds={s['id']:m.parse_source(s,ROOT/manifest[s['id']]['path'],cls.settings,date(2026,10,13)) for s in sources}
        cls.ds['fx_blue']=m.parse_fx(ROOT/manifest['fx']['path'],cls.settings['long_exchange_rate'],date(2026,10,13))
        by_id={s['id']:s for s in sources}
        cls.original={k:m.parse_source(by_id[k],ROOT/path,cls.settings,date(2026,10,13)) for k,path in CONTROL_SNAPSHOTS.items()}
        cls.cards=m.build_charts(cls.ds,cls.settings)

    def test_source_controls(self):
        # Cifras contrastadas con los originales de la primera publicación.
        self.assertAlmostEqual(self.original['icc'][0]['2026-08']['values'][0],34801.42,places=2)
        self.assertAlmostEqual(self.original['rent'][0]['2026-08']['values'][0],597578.88723,places=2)
        self.assertAlmostEqual(self.original['ripte'][0]['2026-07']['values'][0],1946028.12,places=2)
        self.assertAlmostEqual(self.original['new1'][0]['2026-06']['values'][0],3105.25968287032,places=5)
        self.assertAlmostEqual(self.original['cost_m2'][0]['2026-08']['values'][0],1637795.31,places=2)
        # Informe del Colegio de Escribanos de octubre de 2026: 6776 escrituras, $1.200.824 millones, 1208 con hipoteca.
        oct26=self.original['deeds_ce'][0]['2026-10']
        self.assertEqual((oct26['values'][0],oct26['amount'],oct26['mortgages']),(6776,1200824e6,1208))
        # Canasta básica total por adulto equivalente de junio de 2025 (INDEC), citada en el video.
        self.assertAlmostEqual(self.original['cbt'][0]['2025-06']['values'][1],365177.35,places=2)
        # UVA del 31/3/2016 = 14,05 (base de la serie del BCRA).
        self.assertAlmostEqual(self.original['uva'][0]['2016-03']['values'][0],14.05,places=2)
        # La serie actual no puede perder historia respecto de la primera publicación.
        self.assertGreaterEqual(len(self.ds['ripte'][0]),385)

    def test_all_twelve_and_correct_frequency(self):
        self.assertEqual([c['id'] for c in self.cards],list(range(1,13)))
        for c in self.cards:
            for v in c['views']:
                for line in v['lines']:
                    keys=[p[0] for p in line['points']]
                    self.assertEqual(keys,sorted(set(keys)))
                    if v['frequency']=='Trimestral': self.assertTrue(all(k[-2:] in ('03','06','09','12') for k in keys))
        # Antes de 2019 no existe un trimestre completo de MEP en esta fuente.
        self.assertEqual(self.cards[5]['views'][0]['lines'][0]['points'][0][0],'2019-03')

    def test_missing_month_excludes_quarter(self):
        changed=copy.deepcopy(self.ds)
        changed['fx'][0].pop('2025-02')
        cards=m.build_charts(changed,self.settings)
        for card in cards[5:7]:
            for view in card['views']:
                for line in view['lines']:
                    self.assertNotIn('2025-03',[p[0] for p in line['points']])
                    self.assertIn('2025-06',[p[0] for p in line['points']])

    def test_recovery_is_inverse_and_real_rent_base_matches(self):
        returns=self.cards[6]['views'][0]['lines']
        years=self.cards[6]['views'][1]['lines']
        for a,b in zip(returns,years):
            for p,q in zip(a['points'],b['points']):
                self.assertEqual(p[0],q[0]); self.assertAlmostEqual(p[1]*q[1],100,places=10)
        real=self.cards[3]['views'][0]['lines'][0]['points'][-1]
        nominal=self.ds['rent'][0][real[0]]['values'][0]
        self.assertAlmostEqual(real[1],nominal,places=6)
        for line in self.cards[0]['views'][0]['lines']:
            self.assertAlmostEqual(line['points'][0][1],100,places=6)

    def test_salary_conversion_uses_matching_three_months(self):
        # Calcular independientemente el salario en USD de cada mes y su media.
        q='2026-06'; wage=self.ds['ripte'][0]; fx=self.ds['fx'][0]
        usd=(wage['2026-04']['values'][0]/fx['2026-04']['values'][0]+wage['2026-05']['values'][0]/fx['2026-05']['values'][0]+wage['2026-06']['values'][0]/fx['2026-06']['values'][0])/3
        expected=self.ds['used1'][0][q]['values'][0]/usd
        actual=next(p[1] for p in self.cards[5]['views'][0]['lines'][1]['points'] if p[0]==q)
        self.assertAlmostEqual(actual,expected,places=12)

    def test_construction_cost_uses_official_total_and_same_month_mep(self):
        for i,line in enumerate(self.cards[7]['views'][1]['lines']):
            for k,value,provisional in line['points']:
                cost=self.ds['cost_m2'][0][k]
                fx=self.ds['fx'][0][k]
                self.assertAlmostEqual(value,cost['values'][i]/fx['values'][0],places=10)
                self.assertEqual(provisional,cost['provisional'] or fx['provisional'])

    def test_construction_cost_does_not_bridge_missing_exchange_month(self):
        changed=copy.deepcopy(self.ds)
        changed['fx'][0]['2025-02']['observations']=9
        card=m.build_charts(changed,self.settings)[7]
        self.assertNotIn('2025-02',[p[0] for p in card['views'][0]['lines'][0]['points']])
        changes={p[0]:p[1] for p in card['views'][2]['lines'][0]['points']}
        self.assertNotIn('2025-03',changes)
        self.assertIn('2025-04',changes)
        points=dict((p[0],p[1]) for p in card['views'][0]['lines'][0]['points'])
        self.assertAlmostEqual(changes['2025-04'],100*(points['2025-04']/points['2025-03']-1),places=10)

    def test_deeds_prefer_colegio_and_reject_inconsistent_idecba_years(self):
        merged,rejected=m.merge_deeds(self.ds['deeds'][0],self.ds['deeds_ce'][0],self.settings['deeds_tolerance'])
        # IDECBA publicó para 2026 cifras de alrededor de un quinto de las del Colegio.
        self.assertIn('2026',rejected)
        self.assertNotIn('2025',rejected)
        self.assertEqual(merged['2026-10']['source'],'Colegio')
        # Meses sin informe del Colegio se completan con IDECBA (2016-04 no tiene informe publicado).
        self.assertEqual(merged['2016-04']['source'],'IDECBA')
        self.assertEqual(merged['2002-01']['sales'],self.ds['deeds'][0]['2002-01']['values'][0])
        # Un mes de un año descartado sin informe del Colegio queda fuera, no se completa con IDECBA.
        changed=copy.deepcopy(self.ds['deeds_ce'][0]); changed.pop('2026-05')
        merged,_=m.merge_deeds(self.ds['deeds'][0],changed,self.settings['deeds_tolerance'])
        self.assertNotIn('2026-05',merged)

    def test_rolling_average_needs_twelve_consecutive_months(self):
        points=[[f'2020-{i:02d}',float(i),False] for i in range(1,13)]+[['2021-02',1.0,False]]
        out=m.rolling(points)
        self.assertEqual(out,[['2020-12',6.5,False]])

    def test_uva_index_and_payment(self):
        card=self.cards[9]
        uva=self.ds['uva'][0]; base=self.settings['uva_base']
        first=card['views'][1]['lines'][0]['points'][0]
        self.assertEqual(first[0],base); self.assertAlmostEqual(first[1],100,places=10)
        k='2025-06'
        expected=100*uva[k]['values'][0]/uva[base]['values'][0]
        self.assertAlmostEqual(dict((p[0],p[1]) for p in card['views'][1]['lines'][0]['points'])[k],expected,places=10)
        # Índice relativo = índice de la cuota / índice del alquiler de 2 ambientes.
        rent=self.ds['rent'][0]
        rel=expected/(100*rent[k]['values'][1]/rent[base]['values'][1])*100
        self.assertAlmostEqual(dict((p[0],p[1]) for p in card['views'][0]['lines'][0]['points'])[k],rel,places=10)
        # Cuota francesa: 100.000 a 20 años al 4,5% anual = 632,65 por mes.
        self.assertAlmostEqual(m.french_payment(100000,4.5,20),632.65,places=2)
        # Cuota inicial / alquiler, calculada aparte para el 2T 2025.
        months=['2025-04','2025-05','2025-06']; fx=self.ds['fx'][0]; rent=self.ds['rent'][0]
        principal=self.ds['used2'][0][k]['values'][0]*43*sum(fx[x]['values'][0] for x in months)/3*self.settings['mortgage_ltv']
        ratio=100*m.french_payment(principal,4.5,20)/(sum(rent[x]['values'][1] for x in months)/3)
        self.assertAlmostEqual(dict((p[0],p[1]) for p in card['views'][2]['lines'][0]['points'])[k],ratio,places=10)

    def test_salary_and_basket_in_dollars_match_video_figures(self):
        # El video (agosto de 2025) cita RIPTE de junio de 2025 en USD 1238 y la canasta en USD 308.
        mep=self.cards[10]['views'][0]['lines']
        wage=dict((p[0],p[1]) for p in mep[0]['points'])['2025-06']
        basket=dict((p[0],p[1]) for p in mep[1]['points'])['2025-06']
        self.assertAlmostEqual(wage,1238,delta=1238*0.02)
        self.assertAlmostEqual(basket,308,delta=308*0.02)

    def test_yield_by_rooms_uses_matching_rent_and_area(self):
        view=self.cards[6]['views'][2]
        self.assertEqual(view['title'],'Por tipología · usados')
        k='2025-06'; months=['2025-04','2025-05','2025-06']; fx=self.ds['fx'][0]; rent=self.ds['rent'][0]
        for i,area in enumerate(self.settings['rent_areas_m2']):
            monthly=sum(rent[x]['values'][i]/fx[x]['values'][0] for x in months)/3
            expected=100*12*monthly/(self.ds[f'used{i+1}'][0][k]['values'][0]*area)
            self.assertAlmostEqual(dict((p[0],p[1]) for p in view['lines'][i]['points'])[k],expected,places=10)
        # La vista anterior de monoambientes usados coincide con la nueva de 1 ambiente.
        for a,b in zip(self.cards[6]['views'][0]['lines'][1]['points'],view['lines'][0]['points'],strict=True):
            self.assertEqual((a[0],a[2]),(b[0],b[2])); self.assertAlmostEqual(a[1],b[1],places=10)

if __name__=='__main__': unittest.main()
