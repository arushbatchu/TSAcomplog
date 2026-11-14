import sys, os, re, glob
from openpyxl import load_workbook


def find_file(folder, word, required=True):
    for p in sorted(glob.glob(os.path.join(folder, "*.xlsx"))):
        name = os.path.basename(p).lower()
        if word in name and "solution" not in name and not name.startswith("~$"):
            return p
    if required:
        sys.exit(f"no file with '{word}' in its name inside {folder}")
    return None


def rows(path):
    ws = load_workbook(path, data_only=True).worksheets[0]
    out = [list(r) for r in ws.iter_rows(values_only=True)]
    return [r for r in out if any(c is not None and str(c).strip() for c in r)]


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def text(v):
    return str(v).strip() if v is not None else ""


def code(s):
    found = re.findall(r"\(([^)]*)\)", str(s))
    return found[-1].strip().lower() if found else str(s).strip().lower()


class Data:
    def miles(self, a, b):
        if (a, b) in self.distance:
            return self.distance[(a, b)]
        if (b, a) in self.distance:
            return self.distance[(b, a)]
        sys.exit(f"no distance between {a} and {b}")


def load_data(folder):
    d = Data()

    r = rows(find_file(folder, "requirement"))
    d.materials = [text(c) for c in r[0][1:] if text(c)]
    d.products = [text(row[0]) for row in r[1:]]
    d.req = {}
    for row in r[1:]:
        for i, m in enumerate(d.materials, start=1):
            d.req[(text(row[0]), m)] = num(row[i]) or 0.0

    bycode = {code(p): p for p in d.products}

    r = rows(find_file(folder, "supplier"))
    head = [text(c) for c in r[0]]
    cost_cols, cap_cols = {}, {}
    for i, h in enumerate(head):
        who = re.findall(r"\(([^)]*)\)", h)
        if not who:
            continue
        if "cost" in h.lower():
            cost_cols[i] = who[-1].strip()
        elif "max" in h.lower() or "monthly" in h.lower():
            cap_cols[i] = who[-1].strip()
    d.suppliers = list(dict.fromkeys(cost_cols.values()))
    d.cost, d.cap = {}, {}
    for row in r[1:]:
        m = text(row[0])
        for i, s in cost_cols.items():
            if num(row[i]) is not None:
                d.cost[(s, m)] = num(row[i])
        for i, s in cap_cols.items():
            if num(row[i]) is not None:
                d.cap[(s, m)] = num(row[i])

    r = rows(find_file(folder, "processor"))
    head = [text(c) for c in r[0]]
    time_cols, rate_col, cap_col = {}, None, None
    for i, h in enumerate(head):
        low = h.lower()
        if code(h) in bycode:
            time_cols[i] = bycode[code(h)]
        elif "cost" in low:
            rate_col = i
        elif "capacity" in low or "available" in low:
            cap_col = i
    if not time_cols:
        print("  (guessing the processing time columns, double check them)")
        time_cols = {i + 1: p for i, p in enumerate(d.products)}
    d.processors = [text(row[0]) for row in r[1:]]
    d.time, d.rate, d.hours = {}, {}, {}
    for row in r[1:]:
        f = text(row[0])
        for i, p in time_cols.items():
            if num(row[i]) is not None:
                d.time[(f, p)] = num(row[i])
        d.rate[f] = num(row[rate_col]) or 0.0
        d.hours[f] = num(row[cap_col]) or 0.0

    r = rows(find_file(folder, "distributor"))
    head = [text(c) for c in r[0]]
    dem_cols, price_cols = {}, {}
    for i, h in enumerate(head):
        if code(h) in bycode:
            if "demand" in h.lower():
                dem_cols[i] = bycode[code(h)]
            elif "price" in h.lower():
                price_cols[i] = bycode[code(h)]
    if not dem_cols or not price_cols:
        print("  (guessing the demand and price columns, double check them)")
        dem_cols = {1 + 2 * i: p for i, p in enumerate(d.products)}
        price_cols = {2 + 2 * i: p for i, p in enumerate(d.products)}
    d.distributors = [text(row[0]) for row in r[1:]]
    d.demand, d.price = {}, {}
    for row in r[1:]:
        dist = text(row[0])
        for i, p in dem_cols.items():
            d.demand[(dist, p)] = num(row[i]) or 0.0
        for i, p in price_cols.items():
            d.price[(dist, p)] = num(row[i]) or 0.0

    d.ship = {}
    for row in rows(find_file(folder, "shipping"))[1:]:
        vals = [num(c) for c in row[1:] if num(c) is not None]
        if vals:
            d.ship[text(row[0])] = vals[-1]
    for thing in d.materials + d.products:
        if thing not in d.ship:
            match = [v for k, v in d.ship.items() if code(k) == code(thing)]
            if not match:
                sys.exit(f"no shipping cost found for {thing}")
            d.ship[thing] = match[0]

    r = rows(find_file(folder, "distance"))
    head = [text(c) for c in r[0]]
    d.distance = {}
    for row in r[1:]:
        for i in range(1, len(head)):
            if i < len(row) and num(row[i]) is not None:
                d.distance[(text(row[0]), head[i])] = num(row[i])

    d.sup_share = d.sup_off = d.proc_share = d.proc_off = 0.0
    path = find_file(folder, "discount", required=False)
    if path:
        r = rows(path)
        head = [text(c).lower() for c in r[0]]
        who_col = next((i for i, h in enumerate(head) if "applies" in h), 0)
        min_col = next((i for i, h in enumerate(head) if "min" in h or "usage" in h), 1)
        off_col = next((i for i, h in enumerate(head) if "discount" in h), 2)
        for row in r[1:]:
            who = text(row[who_col]).lower()
            share, off = num(row[min_col]), num(row[off_col])
            if share is None or off is None:
                continue
            if "suppl" in who:
                d.sup_share, d.sup_off = share, off
            elif "process" in who or "plant" in who:
                d.proc_share, d.proc_off = share, off
    return d


def solve(d, time_limit=120):
    import numpy as np
    from scipy.optimize import milp, LinearConstraint, Bounds

    buys = [(s, m, f) for s in d.suppliers for m in d.materials for f in d.processors
            if (s, m) in d.cost]
    makes = [(f, p) for f in d.processors for p in d.products if (f, p) in d.time]
    sells = [(f, p, x) for (f, p) in makes for x in d.distributors
             if d.demand.get((x, p), 0) > 0]
    pairs = [(s, m) for s in d.suppliers for m in d.materials if (s, m) in d.cost]

    keys = ([("buy",) + k for k in buys] + [("make",) + k for k in makes] +
            [("sell",) + k for k in sells] +
            [("mcheap",) + k for k in pairs] + [("mfull",) + k for k in pairs] +
            [("mflag",) + k for k in pairs] +
            [("hcheap", f) for f in d.processors] + [("hfull", f) for f in d.processors] +
            [("hflag", f) for f in d.processors])
    at = {k: i for i, k in enumerate(keys)}
    n = len(keys)

    c = np.zeros(n)
    for (s, m, f) in buys:
        c[at[("buy", s, m, f)]] = d.ship[m] * d.miles(s, f)
    for (s, m) in pairs:
        c[at[("mcheap", s, m)]] = d.cost[(s, m)] * (1 - d.sup_off)
        c[at[("mfull", s, m)]] = d.cost[(s, m)]
    for f in d.processors:
        c[at[("hcheap", f)]] = d.rate[f] * (1 - d.proc_off)
        c[at[("hfull", f)]] = d.rate[f]
    for (f, p, x) in sells:
        c[at[("sell", f, p, x)]] = d.ship[p] * d.miles(f, x) - d.price[(x, p)]

    A, lo, hi = [], [], []

    def blank():
        return np.zeros(n)

    for (s, m) in pairs:
        cap = d.cap.get((s, m), 0.0)
        r = blank()
        for f in d.processors:
            if ("buy", s, m, f) in at:
                r[at[("buy", s, m, f)]] = 1
        r[at[("mcheap", s, m)]] = -1
        r[at[("mfull", s, m)]] = -1
        A.append(r); lo.append(0.0); hi.append(0.0)

        r = blank(); r[at[("mcheap", s, m)]] = 1; r[at[("mflag", s, m)]] = -cap
        A.append(r); lo.append(-np.inf); hi.append(0.0)

        r = blank(); r[at[("mcheap", s, m)]] = 1
        r[at[("mflag", s, m)]] = -d.sup_share * cap
        A.append(r); lo.append(0.0); hi.append(np.inf)

        r = blank(); r[at[("mfull", s, m)]] = 1; r[at[("mflag", s, m)]] = cap
        A.append(r); lo.append(-np.inf); hi.append(cap)

        r = blank()
        for f in d.processors:
            if ("buy", s, m, f) in at:
                r[at[("buy", s, m, f)]] = 1
        A.append(r); lo.append(-np.inf); hi.append(cap)

    for f in d.processors:
        cap = d.hours[f]
        r = blank()
        for p in d.products:
            if ("make", f, p) in at:
                r[at[("make", f, p)]] = d.time[(f, p)]
        r[at[("hcheap", f)]] = -1
        r[at[("hfull", f)]] = -1
        A.append(r); lo.append(0.0); hi.append(0.0)

        r = blank(); r[at[("hcheap", f)]] = 1; r[at[("hflag", f)]] = -cap
        A.append(r); lo.append(-np.inf); hi.append(0.0)

        r = blank(); r[at[("hcheap", f)]] = 1
        r[at[("hflag", f)]] = -d.proc_share * cap
        A.append(r); lo.append(0.0); hi.append(np.inf)

        r = blank(); r[at[("hfull", f)]] = 1; r[at[("hflag", f)]] = cap
        A.append(r); lo.append(-np.inf); hi.append(cap)

        r = blank()
        for p in d.products:
            if ("make", f, p) in at:
                r[at[("make", f, p)]] = d.time[(f, p)]
        A.append(r); lo.append(-np.inf); hi.append(cap)

    for f in d.processors:
        for m in d.materials:
            r = blank()
            for s in d.suppliers:
                if ("buy", s, m, f) in at:
                    r[at[("buy", s, m, f)]] = 1
            for p in d.products:
                if ("make", f, p) in at:
                    r[at[("make", f, p)]] -= d.req.get((p, m), 0.0)
            A.append(r); lo.append(0.0); hi.append(np.inf)

    for (f, p) in makes:
        r = blank()
        r[at[("make", f, p)]] = -1
        for x in d.distributors:
            if ("sell", f, p, x) in at:
                r[at[("sell", f, p, x)]] = 1
        A.append(r); lo.append(0.0); hi.append(0.0)

    for x in d.distributors:
        for p in d.products:
            r = blank()
            for f in d.processors:
                if ("sell", f, p, x) in at:
                    r[at[("sell", f, p, x)]] = 1
            A.append(r); lo.append(-np.inf); hi.append(d.demand.get((x, p), 0))

    whole = np.zeros(n)
    for k in makes:
        whole[at[("make",) + k]] = 1
    for k in sells:
        whole[at[("sell",) + k]] = 1
    for k in pairs:
        whole[at[("mflag",) + k]] = 1
    for f in d.processors:
        whole[at[("hflag", f)]] = 1

    upper = np.full(n, np.inf)
    for k in pairs:
        upper[at[("mflag",) + k]] = 1 if d.sup_off > 0 else 0
    for f in d.processors:
        upper[at[("hflag", f)]] = 1 if d.proc_off > 0 else 0

    res = milp(c=c, constraints=LinearConstraint(np.array(A), lo, hi),
               integrality=whole, bounds=Bounds(0, upper),
               options={"time_limit": time_limit})
    if res.x is None:
        sys.exit(f"solver failed: {res.message}")

    buy = {k: round(res.x[at[("buy",) + k]], 1) for k in buys
           if res.x[at[("buy",) + k]] > 1e-6}
    make = {k: int(round(res.x[at[("make",) + k]])) for k in makes
            if res.x[at[("make",) + k]] > 1e-6}
    sell = {k: int(round(res.x[at[("sell",) + k]])) for k in sells
            if res.x[at[("sell",) + k]] > 1e-6}
    return buy, make, sell


def check(d, buy, make, sell, quiet=False):
    revenue = material = inbound = processing = outbound = 0.0
    problems = []
    got_sup, got_proc = [], []

    for (s, m) in sorted({(s, m) for (s, m, f) in buy}):
        total = sum(q for (ss, mm, _), q in buy.items() if ss == s and mm == m)
        cap = d.cap.get((s, m), 0.0)
        earned = d.sup_off > 0 and cap > 0 and total >= d.sup_share * cap - 1e-6
        material += d.cost[(s, m)] * (1 - d.sup_off if earned else 1) * total
        if earned:
            got_sup.append(f"{s} / {m}: {total:g} of {cap:g}")
        if total > cap + 1e-6:
            problems.append(f"{s} only has {cap} {m} but the plan buys {total}")

    for (s, m, f), q in buy.items():
        inbound += d.ship[m] * d.miles(s, f) * q

    for f in d.processors:
        hours = sum(d.time[(ff, p)] * u for (ff, p), u in make.items() if ff == f)
        cap = d.hours[f]
        earned = d.proc_off > 0 and cap > 0 and hours >= d.proc_share * cap - 1e-6
        processing += d.rate[f] * (1 - d.proc_off if earned else 1) * hours
        if earned:
            got_proc.append(f"{f}: {hours:g} h of {cap:g}")
        if hours > cap + 1e-6:
            problems.append(f"{f} uses {hours:.1f} hours but only has {cap}")
        for m in d.materials:
            held = sum(q for (_, mm, ff), q in buy.items() if mm == m and ff == f)
            need = sum(d.req.get((p, m), 0.0) * u for (ff, p), u in make.items() if ff == f)
            if need > held + 1e-4:
                problems.append(f"{f} is short {need-held:.2f} of {m}")

    for (f, p, x), u in sell.items():
        revenue += d.price[(x, p)] * u
        outbound += d.ship[p] * d.miles(f, x) * u

    for (f, p), u in make.items():
        shipped = sum(v for (ff, pp, _), v in sell.items() if ff == f and pp == p)
        if shipped != u:
            problems.append(f"{f} builds {u} {p} but ships {shipped}")
    for x in d.distributors:
        for p in d.products:
            v = sum(u for (_, pp, xx), u in sell.items() if pp == p and xx == x)
            if v > d.demand.get((x, p), 0) + 1e-6:
                problems.append(f"{x} only wants {d.demand.get((x,p),0)} {p}, plan sends {v}")

    profit = revenue - material - inbound - processing - outbound
    if not quiet:
        if got_sup or got_proc:
            print("  discounts:")
            for discount in got_sup + got_proc:
                print("   ", discount)
        print(f"  revenue            {revenue:>12,.2f}")
        print(f"  material cost     -{material:>12,.2f}")
        print(f"  shipping in       -{inbound:>12,.2f}")
        print(f"  processing cost   -{processing:>12,.2f}")
        print(f"  shipping out      -{outbound:>12,.2f}")
        print(f"  PROFIT             {profit:>12,.2f}")
        for problem in problems:
            print("  Issue:", problem)
    return profit, not problems


def parts(label):
    label = str(label).replace("->", "→")
    left, _, right = label.rpartition("\\")
    out = [p.strip() for p in left.split("→") if p.strip()]
    if right.strip():
        out.append(right.strip())
    return out


def find_template(folder):
    here = glob.glob(os.path.join(folder, "*template*.xlsx"))
    if here:
        return here[0]
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    there = glob.glob(os.path.join(root, "config", "*template*.xlsx"))
    return there[0] if there else None


def write(d, buy, make, sell, template, out_path):
    wb = load_workbook(template)
    ws = wb.worksheets[0]
    missed = 0

    kind_col, name_col, qty_col = 1, 2, 3
    for c in range(1, ws.max_column + 1):
        head = text(ws.cell(row=1, column=c).value).lower()
        if "type" in head:
            kind_col = c
        elif "compan" in head or "name" in head:
            name_col = c
        elif "quantity" in head or "amount" in head:
            qty_col = c

    for row in range(2, ws.max_row + 1):
        kind = text(ws.cell(row=row, column=kind_col).value).lower()
        bits = parts(ws.cell(row=row, column=name_col).value)
        if not kind or not bits:
            continue
        value = None
        if "distrib" in kind and len(bits) == 3:
            value = sell.get((bits[0], bits[2], bits[1]), 0)
        elif "ship" in kind and len(bits) == 3:
            value = buy.get((bits[0], bits[2], bits[1]), 0)
        elif "process" in kind and len(bits) == 2:
            value = make.get((bits[0], bits[1]), 0)
        elif len(bits) == 2:
            value = round(sum(q for (s, m, _), q in buy.items()
                              if s == bits[0] and m == bits[1]), 1)
        if value is None:
            missed += 1
            print(f"  !! row {row} did not match anything: {bits}")
        else:
            ws.cell(row=row, column=qty_col).value = value

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    try:
        wb.save(out_path)
    except PermissionError:
        sys.exit(f"could not write {out_path}, close it in Excel and run again")
    return missed == 0


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else "config"
    d = load_data(folder)

    buy, make, sell = solve(d)

    print("Plan:")
    for (s, m, f), q in sorted(buy.items()):
        print(f"  buy  {q:>9.1f}  {m:<28} {s:<16} -> {f}")
    for (f, p), u in sorted(make.items()):
        print(f"  make {u:>9}  {p:<28} at {f} ({d.time[(f,p)]*u:.1f} h)")
    for (f, p, x), u in sorted(sell.items()):
        print(f"  ship {u:>9}  {p:<28} {f:<22} -> {x}")

    print("\nCosts:")
    profit, ok = check(d, buy, make, sell)

    template = find_template(folder)
    if not template:
        print("No solution template found.")
        return 1
    wrote = write(d, buy, make, sell, template, "out/solution.xlsx")
    print("\nSaved: out/solution.xlsx")
    if not ok or not wrote:
        print("Review the issues above before submitting.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
