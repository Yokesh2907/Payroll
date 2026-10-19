"""Run with:  python3 -m unittest -v"""
import unittest
from payroll_calc import compute, num, D

def emp(**kw):
    base = dict(working_days='25', paid_holidays='1', basic='0', allowances='0', other_earnings='0', ot_method='Hours × OT Rate',
                multiplier='1', hours_per_day='8', ot_hours='0', ot_rate='0', festival_recovery='0', general_recovery='0',
                general_advance='0', previous_balance='0', other_deductions='0', esi_exempt='No')
    base.update(kw); return base

class TestWorkbookEmployees(unittest.TestCase):
    """Expected figures come from the supplied salary sheets."""
    def test_balaji_matches_his_sheet(self):
        c, w = compute(emp(basic='12012', allowances='2002', other_earnings='5980', ot_hours='10', ot_rate='144.20',
                           festival_recovery='200', general_recovery='300', previous_balance='1000'))
        self.assertEqual(c['ot_amount'], '1442.00')      # sheet: OT 1442
        self.assertEqual(c['pf'], '1441.00')             # sheet: EPF 1441
        self.assertEqual(c['esi'], '135.00')             # sheet: ESI 135
        self.assertEqual(c['gross'], '21436.00')
        self.assertEqual(c['net_salary'], '19360.00')
        self.assertEqual(num(c['epf']) + num(c['eps']), num(c['pf']))
        self.assertEqual(w, [])

    def test_govindaraj_esi_exempt_and_above_ceiling(self):
        c, _ = compute(emp(basic='13350', allowances='2670', other_earnings='10680', ot_hours='10', ot_rate='102.70', esi_exempt='Yes'))
        self.assertEqual(c['esi'], '0.00')
        c, _ = compute(emp(basic='13350', allowances='2670', other_earnings='10680'))   # not flagged, but gross 26700 > 21000
        self.assertEqual(c['esi'], '0.00')

    def test_arman_uses_monthly_basic(self):
        c, _ = compute(emp(basic='11700', allowances='1950', other_earnings='5850', ot_hours='10', ot_rate='90'))
        self.assertEqual(c['gross'], '20400.00')
        self.assertEqual(c['esi'], '132.00')

class TestRules(unittest.TestCase):
    def test_blank_inputs_do_not_crash(self):
        c, _ = compute(emp(basic='', ot_hours='', allowances=None))
        self.assertEqual(c['net_salary'], '0.00')

    def test_invalid_text_raises(self):
        with self.assertRaises(ValueError): compute(emp(basic='abc'))

    def test_ot_hours_without_rate_warns(self):
        _, w = compute(emp(basic='1000', ot_hours='10', ot_rate='0'))
        self.assertTrue(any('OT' in m for _, m, _ in w))

    def test_ot_from_basic_method(self):
        c, _ = compute(emp(basic='12012', ot_hours='10', multiplier='1.5',
                           ot_method='Basic ÷ Working Days ÷ Hours Per Day × OT Hours × Multiplier'))
        self.assertEqual(c['ot_amount'], '866.25')

    def test_negative_net_is_an_error(self):
        _, w = compute(emp(basic='1000', other_deductions='999999'))
        self.assertTrue(any(l == 'error' for l, _, _ in w))

    def test_proration(self):
        c, _ = compute(emp(basic='13000', working_days='19', paid_holidays='1'), {'company_days': D('26')})   # 20/26 paid
        self.assertEqual(c['gross'], '10000.00'); self.assertEqual(c['paid_days'], '20')   # whole number, no decimals
        c, _ = compute(emp(basic='13000', working_days='19', paid_holidays='1'), {'company_days': D('26'), 'prorate': False})
        self.assertEqual(c['gross'], '13000.00')

    def test_pf_ceiling(self):
        c, _ = compute(emp(basic='30000'));                          self.assertEqual(c['pf'], '1800.00')   # 12% of 15000
        c, _ = compute(emp(basic='30000'), {'pf_ceiling': D('0')});  self.assertEqual(c['pf'], '3600.00')

    def test_esi_threshold_ignores_overtime(self):
        c, _ = compute(emp(basic='15000', allowances='2000', other_earnings='3000', ot_hours='10', ot_rate='200'))  # wages 20000 + OT 2000
        self.assertNotEqual(c['esi'], '0.00')

    def test_deferred_general_recovery_does_not_reduce_current_salary(self):
        normal, _ = compute(emp(basic='10000', general_advance='1000', general_recovery='300', previous_balance='500'))
        deferred, _ = compute(emp(basic='10000', general_advance='1000', general_recovery='300', previous_balance='500', defer_general_recovery='Yes'))
        self.assertEqual(num(deferred['net_salary']) - num(normal['net_salary']), D('300.00'))
        self.assertEqual(deferred['deferred_recovery'], '300.00')
        self.assertEqual(deferred['general_balance'], '1500.00')

    def test_festival_advance_defer_and_balance(self):
        normal, _ = compute(emp(basic='10000', festival_advance='1000', festival_recovery='250', festival_previous_balance='500'))
        deferred, _ = compute(emp(basic='10000', festival_advance='1000', festival_recovery='250', festival_previous_balance='500',
                                  defer_festival_recovery='Yes'))
        self.assertEqual(normal['festival_balance'], '1250.00')
        self.assertEqual(deferred['festival_balance'], '1500.00')
        self.assertEqual(deferred['deferred_festival_recovery'], '250.00')
        self.assertEqual(num(deferred['net_salary']) - num(normal['net_salary']), D('250.00'))
        self.assertEqual(normal['advance_deduction'], '250.00'); self.assertEqual(deferred['advance_deduction'], '0.00')

    def test_advance_totals(self):
        c, _ = compute(emp(basic='10000', general_recovery='100', previous_balance='400', festival_recovery='50', festival_previous_balance='200'))
        self.assertEqual(c['advance_deduction'], '150.00'); self.assertEqual(c['advance_balance'], '450.00')

    def test_fractional_days_warn(self):
        _, w = compute(emp(working_days='24.5', basic='1000'))
        self.assertTrue(any('whole' in m for _, m, _ in w))

    def test_advance_balance_not_clamped(self):
        c, w = compute(emp(previous_balance='100', general_recovery='300'))
        self.assertEqual(c['general_balance'], '-200.00'); self.assertTrue(any('advance' in m for _, m, _ in w))

if __name__ == '__main__':
    unittest.main()
