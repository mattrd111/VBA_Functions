"""Checks the fund maths used by src/modFinance.bas and src/modWaterfall.bas.

VBA cannot be run outside Excel, so this is a line-by-line mirror of the
algorithms in Python, checked against answers that are known analytically
(-100 today and +110 in a year is 10%, and no solver gets a vote on that).

It verifies the ALGORITHMS, not the VBA implementation - the two have to be
kept in step by hand. If you change the solver, the preferred-return walk or
the waterfall tiers, change them here too and re-run:

    python build/test_fund_maths.py

The strongest test here is the last one: 20,000 random waterfalls, checking
that the tiers always sum back to the amount distributed and that the GP never
takes more than its carry share of the profit. A waterfall that does not tie
out is the kind of error that survives a review.
"""

import random

DAYS = 365.0

# ---------------------------------------------------------------- XNPV / XIRR
def xnpv(rate, values, dates):
    d0 = dates[0]
    if rate <= -1.0:
        return None
    return sum(v / (1.0 + rate) ** ((d - d0) / DAYS) for v, d in zip(values, dates))

def xirr(values, dates):
    """Bracket then bisect. Always converges when a root exists, unlike Newton
    from a fixed guess, which is what makes Excel's XIRR give up."""
    if len(values) < 2:
        return ('#NUM', 'need at least two cash flows')
    if not (any(v > 0 for v in values) and any(v < 0 for v in values)):
        return ('#NUM', 'cash flows must change sign')

    lo, hi = -0.9999, -0.9999
    f_lo = xnpv(lo, values, dates)
    # walk the rate up until the sign of the NPV flips
    rate, step, found = -0.99, 0.01, False
    while rate < 1000.0:
        f = xnpv(rate, values, dates)
        if f is not None and f_lo is not None and f * f_lo <= 0:
            hi, f_hi = rate, f
            found = True
            break
        lo, f_lo = rate, f
        rate += step
        step *= 1.05                      # coarser as we go, so 1000% is reachable
    if not found:
        return ('#NUM', 'no rate found')

    for _ in range(200):
        mid = (lo + hi) / 2.0
        f_mid = xnpv(mid, values, dates)
        if abs(f_mid) < 1e-9 or (hi - lo) < 1e-12:
            return mid
        if f_mid * f_lo <= 0:
            hi = mid
        else:
            lo, f_lo = mid, f_mid
    return (lo + hi) / 2.0

# ------------------------------------------------------------------ multiples
def paid_in(values):      return -sum(v for v in values if v < 0)
def distributed(values):  return sum(v for v in values if v > 0)
def dpi(values):          return distributed(values) / paid_in(values) if paid_in(values) else None
def rvpi(values, nav):    return nav / paid_in(values) if paid_in(values) else None
def tvpi(values, nav=0):  return (distributed(values) + nav) / paid_in(values) if paid_in(values) else None

# ------------------------------------------------------------- preferred return
def accrue(balance, days, rate, compound):
    if compound:
        return balance * ((1.0 + rate) ** (days / DAYS) - 1.0)
    return balance * rate * days / DAYS

def pref_state(values, dates, rate, as_of, capital_first=True, compound=True):
    """Returns (unreturned capital, accrued but unpaid preferred) at as_of."""
    order = sorted(range(len(values)), key=lambda i: dates[i])
    unreturned = 0.0
    accrued = 0.0
    last = dates[order[0]]

    for i in order:
        d, a = dates[i], values[i]
        if d > as_of:
            break
        base = unreturned + (accrued if compound else 0.0)
        accrued += accrue(base, d - last, rate, compound) if compound else accrue(unreturned, d - last, rate, False)
        last = d
        if a < 0:
            unreturned += -a
        elif a > 0:
            left = a
            if capital_first:
                p = min(left, unreturned); unreturned -= p; left -= p
                p = min(left, accrued);    accrued    -= p; left -= p
            else:
                p = min(left, accrued);    accrued    -= p; left -= p
                p = min(left, unreturned); unreturned -= p; left -= p

    if as_of > last:
        base = unreturned + (accrued if compound else 0.0)
        accrued += accrue(base, as_of - last, rate, compound) if compound else accrue(unreturned, as_of - last, rate, False)
    return unreturned, accrued

# ----------------------------------------------------------------- waterfall
PARTS = ("LPCAPITAL","LPPREF","GPCATCHUP","LPCATCHUP","LPRESIDUAL","GPCARRY","LPTOTAL","GPTOTAL","TOTAL")

def waterfall(distributable, capital_due, pref_due, carry_rate, catch_up_rate=1.0):
    if carry_rate < 0 or carry_rate >= 1: return None
    if catch_up_rate > 1: catch_up_rate = 1.0
    if catch_up_rate < 0: catch_up_rate = 0.0
    if distributable < 0: distributable = 0.0
    if capital_due < 0: capital_due = 0.0
    if pref_due < 0: pref_due = 0.0

    remaining = distributable
    t1 = min(remaining, max(capital_due, 0.0)); remaining -= t1          # return of capital
    t2 = min(remaining, max(pref_due, 0.0));    remaining -= t2          # preferred

    gp_catch = lp_catch = 0.0
    if catch_up_rate > 0 and carry_rate > 0 and t2 > 0:
        target = carry_rate * t2 / (1.0 - carry_rate)   # GP's share of profit so far
        tier = min(remaining, target / catch_up_rate)
        gp_catch = tier * catch_up_rate
        lp_catch = tier * (1.0 - catch_up_rate)
        remaining -= tier

    lp_res = remaining * (1.0 - carry_rate)
    gp_res = remaining * carry_rate

    return {"LPCAPITAL":t1, "LPPREF":t2, "GPCATCHUP":gp_catch, "LPCATCHUP":lp_catch,
            "LPRESIDUAL":lp_res, "GPCARRY":gp_res,
            "LPTOTAL":t1+t2+lp_catch+lp_res, "GPTOTAL":gp_catch+gp_res,
            "TOTAL":distributable}

# --------------------------------------------------------------------- KS-PME
def kspme(values, dates, index_levels, nav=0.0):
    end = max(index_levels[i] for i in range(len(dates)) if dates[i] == max(dates))
    fv_dist = fv_cont = 0.0
    for v, ix in zip(values, index_levels):
        if ix <= 0: return None
        if v > 0: fv_dist += v * (end / ix)
        elif v < 0: fv_cont += -v * (end / ix)
    if fv_cont == 0: return None
    return (fv_dist + nav) / fv_cont

# ==============================================================================
# Tests
# ==============================================================================
fails = []
def check(name, got, want, tol=1e-6):
    ok = (got is not None and want is not None and abs(got - want) <= tol)
    print(f"{'ok  ' if ok else 'FAIL'} {name:52s} {got if got is None else round(got,8)}  want {want}")
    if not ok: fails.append(name)

print("--- XIRR against analytically known answers ---")
check("-100 -> +110 over 365d",            xirr([-100,110],[0,365]), 0.10)
check("-100 -> +121 over 730d",            xirr([-100,121],[0,730]), 0.10)
check("-100 -> +100 (no gain)",            xirr([-100,100],[0,365]), 0.0)
check("-100 -> +50 (loss)",                xirr([-100,50],[0,365]), -0.50)
check("-1000 -> +2000 over 1095d",         xirr([-1000,2000],[0,1095]), 2**(1/3)-1)
check("multi-flow, NPV must be ~0",        xnpv(xirr([-100,-50,30,40,120],[0,200,500,700,1100]),
                                                [-100,-50,30,40,120],[0,200,500,700,1100]), 0.0, 1e-6)
check("high return -100 -> +500 in 365d",  xirr([-100,500],[0,365]), 4.0)
r = xirr([-100,110],[0,365]); check("stable on repeat",  r, 0.10)
print(f"ok   no sign change returns error: {xirr([-100,-50],[0,365])[0]}")

print("\n--- multiples ---")
cf = [-100,-100,50,80,60]
check("paid in",   paid_in(cf), 200.0)
check("distributed", distributed(cf), 190.0)
check("DPI",       dpi(cf), 0.95)
check("TVPI w/ NAV 60", tvpi(cf,60), 1.25)
check("RVPI w/ NAV 60", rvpi(cf,60), 0.30)

print("\n--- preferred return accrual ---")
u,a = pref_state([-100],[0],0.08,365); check("8% compound, 1yr accrued", a, 8.0)
u,a = pref_state([-100],[0],0.08,730); check("8% compound, 2yr accrued", a, 16.64, 1e-9)
u,a = pref_state([-100],[0],0.08,730,compound=False); check("8% simple, 2yr accrued", a, 16.0)
# accruing in two steps must equal accruing in one
u1,a1 = pref_state([-100,0],[0,365],0.08,730)
check("two-step == one-step",     a1, 16.64, 1e-9)
u,a = pref_state([-100,108],[0,365],0.08,365,capital_first=False)
check("pref-first: 108 clears 8 pref then 100 capital -> 0 accrued", a, 0.0)
check("...and 0 capital left", u, 0.0)
u,a = pref_state([-100,108],[0,365],0.08,365,capital_first=True)
check("capital-first: same totals here", a, 0.0)
u,a = pref_state([-100,50],[0,365],0.08,365,capital_first=True)
check("capital-first: 50 repays capital, pref untouched", a, 8.0)
check("...50 capital left", u, 50.0)

print("\n--- waterfall ---")
w = waterfall(300, 100, 8, 0.20, 1.0)
check("tier 1 return of capital", w["LPCAPITAL"], 100.0)
check("tier 2 preferred",         w["LPPREF"], 8.0)
check("tier 3 GP catch-up",       w["GPCATCHUP"], 2.0)          # 20% of (8+2)=10 profit
check("GP total",                 w["GPTOTAL"], 2.0 + (300-110)*0.20)
check("LP total",                 w["LPTOTAL"], 300 - w["GPTOTAL"])
w = waterfall(105, 100, 8, 0.20, 1.0)
check("short of pref: LP gets all", w["LPTOTAL"], 105.0)
check("short of pref: GP gets none", w["GPTOTAL"], 0.0)
w = waterfall(300, 100, 8, 0.20, 0.5)
check("50% catch-up: GP catch-up still 2", w["GPCATCHUP"], 2.0)
check("50% catch-up: LP gets the other half of that tier", w["LPCATCHUP"], 2.0)
w = waterfall(1000, 100, 8, 0.0, 1.0)
check("zero carry: LP takes everything", w["LPTOTAL"], 1000.0)

print("\n--- waterfall invariant over 20,000 random cases ---")
random.seed(7)
worst = 0.0
for _ in range(20000):
    d  = random.uniform(0, 5000)
    c  = random.uniform(0, 2000)
    p  = random.uniform(0, 500)
    cr = random.choice([0.0, 0.05, 0.10, 0.20, 0.25, 0.30])
    cu = random.choice([0.0, 0.5, 0.8, 1.0])
    w = waterfall(d, c, p, cr, cu)
    worst = max(worst, abs((w["LPTOTAL"] + w["GPTOTAL"]) - d))
    if w["GPTOTAL"] < -1e-9 or w["LPTOTAL"] < -1e-9: fails.append("negative split")
    # the GP can never take more than the carry rate of total profit
    profit = max(d - c, 0.0)
    if profit > 0 and w["GPTOTAL"] > profit * cr + 1e-9: fails.append("GP over carry share")
print(f"{'ok  ' if worst < 1e-9 else 'FAIL'} tiers always sum to the distributable amount (worst drift {worst:.2e})")
if worst >= 1e-9: fails.append("sum invariant")

print("\n--- KS-PME ---")
check("fund matches index exactly -> 1.0", kspme([-100,150],[0,730],[100,150]), 1.0)
check("fund beats index",                  kspme([-100,200],[0,730],[100,150]), 200/(100*1.5))
check("fund lags index",                   kspme([-100,120],[0,730],[100,150]), 120/150)

print()
print(f"{'ALL PASSED' if not fails else 'FAILURES: ' + str(sorted(set(fails)))}")
