"""Comprobaciones de cobertura, cifras de fuente y cálculos temporales."""
import copy
import importlib.util
import json
import unittest
from datetime import date
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('monitor',ROOT/'scripts/update.py')
m=importlib.util.module_from_spec(spec); spec.loader.exec_module(m)

class PipelineTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.settings=json.loads((ROOT/'config/settings.json').read_text())
        sources=json.loads((ROOT/'config/sources.json').read_text())
        manifest=json.loads((ROOT/'data/manifest.json').read_text())
        cls.ds={s['id']:m.parse_source(s,ROOT/manifest[s['id']]['path'],cls.settings,date(2026,10,13)) for s in sources}
        cls.cards=m.build_charts(cls.ds,cls.settings)

    def test_source_controls(self):
        # Cifras contrastadas con los originales de la primera publicación.
        self.assertAlmostEqual(self.ds['icc'][0]['2026-08']['values'][0],34801.42,places=2)
        self.assertAlmostEqual(self.ds['rent'][0]['2026-08']['values'][0],597578.88723,places=2)
        self.assertAlmostEqual(self.ds['ripte'][0]['2026-07']['values'][0],1946028.12,places=2)
        self.assertGreaterEqual(len(self.ds['ripte'][0]),385)
        self.assertAlmostEqual(self.ds['new1'][0]['2026-06']['values'][0],3105.25968287032,places=5)

    def test_all_eight_and_correct_frequency(self):
        self.assertEqual([c['id'] for c in self.cards],list(range(1,9)))
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
        self.assertAlmostEqual(self.ds['cost_m2'][0]['2026-08']['values'][0],1637795.31,places=2)
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

if __name__=='__main__': unittest.main()
