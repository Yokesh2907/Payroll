"""Pure payroll calculation logic (no GUI, no file I/O) so it can be unit-tested.

All inputs are strings/numbers as typed in the grid; all outputs are 2-decimal strings.
"""
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

D = Decimal

OT_METHODS = [
    'Hours × OT Rate',
    'Hours × OT Rate × Multiplier',
    'Basic ÷ Working Days ÷ Hours Per Day × OT Hours × Multiplier',
]
# Older files may contain this value; it was identical to the first method.
LEGACY_OT_METHODS = {'Custom Hourly Rate × OT Hours': OT_METHODS[0]}

CALC_FIELDS = ['paid_days', 'ot_amount', 'gross', 'general_balance', 'pf', 'epf', 'eps',
               'esi', 'total_deductions', 'net_salary', 'deferred_recovery',
               'festival_balance', 'deferred_festival_recovery', 'advance_deduction', 'advance_balance']
WHOLE_FIELDS = ('paid_days',)    # shown as whole numbers (no decimals)

DEFAULT_SETTINGS = {
    'company_days': D('26'),     # divisor for proration and the OT-from-basic method
    'prorate': True,             # pay (working days + paid holidays) / company days
    'pf_rate': D('12'),          # % of basic (employee share)
    'eps_rate': D('8.33'),       # % of basic - reporting split of the PF amount
    'pf_ceiling': D('15000'),    # PF wage ceiling on earned basic; 0 = no ceiling
    'esi_rate': D('0.75'),       # % employee share
    'esi_ceiling': D('21000'),   # ESI applies only if wages (gross excl. OT) <= ceiling; 0 = no ceiling
    'esi_excl_allow': True,      # match workbook: ESI base = gross - OT - washing allowance
    'round_whole': True,         # round deductions to whole rupees (matches the supplied sheets)
}


def num(x):
    """Parse a cell value to Decimal. Blank -> 0. Raises ValueError on junk."""
    if x is None:
        return D('0')
    if isinstance(x, Decimal):
        d = x
    else:
        s = str(x).replace(',', '').replace('₹', '').strip()
        if s == '':
            return D('0')
        try:
            d = D(s)
        except InvalidOperation:
            raise ValueError(f'Invalid numeric value: {x!r}')
    if not d.is_finite():
        raise ValueError(f'Invalid numeric value: {x!r}')
    return d


def num_or_zero(x):
    try:
        return num(x)
    except ValueError:
        return D('0')


def q2(x):
    return x.quantize(D('0.01'), rounding=ROUND_HALF_UP)


def fmt(x):
    return f'{q2(x)}'


def fmt_whole(x):
    return f'{x.quantize(D("1"), rounding=ROUND_HALF_UP)}'


def is_yes(x):
    return str(x).strip().lower() in ('yes', 'y', 'true', '1')


def compute(v, s=None):
    """Compute one employee's payroll.

    v: dict of input values (strings). s: settings dict (see DEFAULT_SETTINGS).
    Returns (calc, warnings): calc maps CALC_FIELDS -> 2-decimal strings;
    warnings is a list of (level, message, field) with level 'error' or 'warn'.
    Raises ValueError for non-numeric input or a zero divisor that the chosen method needs.
    """
    s = {**DEFAULT_SETTINGS, **(s or {})}
    g = lambda k: v.get(k, '')
    basic, allow, other = num(g('basic')), num(g('allowances')), num(g('other_earnings'))
    hrs, rate, mult, hpd = num(g('ot_hours')), num(g('ot_rate')), num(g('multiplier')), num(g('hours_per_day'))
    wd, ph = num(g('working_days')), num(g('paid_holidays'))
    requested_fest_rec, requested_gen_rec, other_ded = num(g('festival_recovery')), num(g('general_recovery')), num(g('other_deductions'))
    # "Defer" = do not deduct this month; the amount moves to next month.
    defer_gen, defer_fest = is_yes(g('defer_general_recovery')), is_yes(g('defer_festival_recovery'))
    gen_rec = D('0') if defer_gen else requested_gen_rec
    deferred_recovery = requested_gen_rec if defer_gen else D('0')
    fest_rec = D('0') if defer_fest else requested_fest_rec
    deferred_fest = requested_fest_rec if defer_fest else D('0')
    gen_given, prev_bal = num(g('general_advance')), num(g('previous_balance'))
    fest_given, fest_prev = num(g('festival_advance')), num(g('festival_previous_balance'))
    cdays = s['company_days']
    exempt = str(g('esi_exempt')).strip().lower() in ('yes', 'y', 'true', '1')
    rnd = (lambda x: x.quantize(D('1'), rounding=ROUND_HALF_UP)) if s['round_whole'] else q2
    warns = []

    # --- attendance / proration -------------------------------------------------
    if wd != wd.to_integral_value() or ph != ph.to_integral_value():
        warns.append(('warn', 'Working days / paid holidays should be whole numbers.', 'paid_days'))
    paid_days = wd + ph
    if cdays > 0 and paid_days > cdays:
        warns.append(('warn', f'Working days + paid holidays ({paid_days}) exceed company working days ({cdays}); capped.', 'paid_days'))
        paid_days = cdays
    factor = (paid_days / cdays) if (s['prorate'] and cdays > 0) else D('1')
    e_basic, e_allow, e_other = basic * factor, allow * factor, other * factor

    # --- overtime ---------------------------------------------------------------
    method = LEGACY_OT_METHODS.get(g('ot_method'), g('ot_method')) or OT_METHODS[0]
    if method == OT_METHODS[1]:
        ot = hrs * rate * mult
    elif method == OT_METHODS[2]:
        if cdays <= 0 or hpd <= 0:
            raise ValueError('Company Working Days and Hours/Day must be greater than zero for this OT method.')
        ot = (basic / cdays / hpd) * hrs * mult
    else:
        ot = hrs * rate
    if hrs > 0 and ot == 0:
        warns.append(('warn', 'OT hours entered but OT amount is 0 - check OT Rate / multiplier.', 'ot_amount'))

    gross = e_basic + e_allow + e_other + ot

    # --- statutory deductions ---------------------------------------------------
    pf_base = e_basic if s['pf_ceiling'] <= 0 else min(e_basic, s['pf_ceiling'])
    pf = rnd(pf_base * s['pf_rate'] / 100)
    eps = min(rnd(pf_base * s['eps_rate'] / 100), pf)
    epf = pf - eps                                   # EPF + EPS always add back to PF

    esi_wages = gross - ot                           # overtime is not ESI wages
    eligible = (not exempt) and (s['esi_ceiling'] <= 0 or esi_wages <= s['esi_ceiling'])
    esi_base = esi_wages - (e_allow if s['esi_excl_allow'] else 0)
    esi = rnd(max(esi_base, D('0')) * s['esi_rate'] / 100) if eligible else D('0')

    total = pf + esi + fest_rec + gen_rec + other_ded
    net = gross - total
    balance = prev_bal + gen_given - gen_rec
    fest_balance = fest_prev + fest_given - fest_rec

    # --- sanity checks ----------------------------------------------------------
    if net < 0:
        warns.append(('error', f'Net salary is negative ({fmt(net)}).', 'net_salary'))
    if balance < 0:
        warns.append(('warn', f'General advance balance is negative ({fmt(balance)}): recovery exceeds balance.', 'general_balance'))
    if fest_balance < 0 and (fest_prev or fest_given):   # only when festival advances are being tracked
        warns.append(('warn', f'Festival advance balance is negative ({fmt(fest_balance)}): recovery exceeds balance.', 'festival_balance'))
    for label, val in [('Basic', basic), ('Allowances', allow), ('Other earnings', other), ('OT hours', hrs),
                       ('Festival recovery', fest_rec), ('General recovery', gen_rec), ('Other deductions', other_ded),
                       ('Working days', wd), ('Paid holidays', ph)]:
        if val < 0:
            warns.append(('error', f'{label} is negative.', 'net_salary'))

    calc = {'paid_days': paid_days, 'festival_balance': fest_balance, 'deferred_festival_recovery': deferred_fest,
            'advance_deduction': fest_rec + gen_rec, 'advance_balance': balance + fest_balance, 'ot_amount': ot, 'gross': gross, 'general_balance': balance, 'pf': pf, 'epf': epf,
            'eps': eps, 'esi': esi, 'total_deductions': total, 'net_salary': net, 'deferred_recovery': deferred_recovery}
    return {k: (fmt_whole(x) if k in WHOLE_FIELDS else fmt(x)) for k, x in calc.items()}, warns
