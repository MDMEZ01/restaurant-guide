import openpyxl, re, json, sys

SRC = sys.argv[1]
OUT = sys.argv[2]

DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']
DI = {d: i for i, d in enumerate(DAYS)}

# Hand-normalized versions of the irregular hours strings
OVERRIDES = {
    'Daily 11 AM to 8:30 or 9 PM (listings vary — confirm)': 'Daily 11 AM–9 PM',
    'Daily 11 AM–9 PM per sweetgreen site (other listings show 10 AM–10 PM)': 'Daily 11 AM–9 PM',
    'Daily 5:30–10 PM (Fri–Sat until 11 PM)': 'Sun–Thu 5:30–10 PM; Fri–Sat 5:30–11 PM',
    'Daily 5–9:30 PM (Fri–Sat until 10 PM)': 'Sun–Thu 5–9:30 PM; Fri–Sat 5–10 PM',
    'Daily 7 AM–9 PM (Fri–Sat until 10 PM)': 'Sun–Thu 7 AM–9 PM; Fri–Sat 7 AM–10 PM',
    'Daily noon–10 PM (Fri–Sat noon–11 PM)': 'Sun–Thu noon–10 PM; Fri–Sat noon–11 PM',
    'Mon & Thu–Sun 5–10 PM (Fri–Sat until 10:30 PM)': 'Mon & Thu & Sun 5–10 PM; Fri–Sat 5–10:30 PM',
    'Daily lunch 11:30 AM–3:30 PM, dinner 5–10:30 PM; delivery via Alimentari next door ~11:30 AM–8 PM': 'Daily 11:30 AM–3:30 PM & 5–10:30 PM',
    'Daily, opening 10–11 AM and closing 7–8 PM (listings vary — confirm)': 'Daily 11 AM–7 PM',
    'Sun–Tue 11:30 AM–11 PM; Wed–Thu to 12 AM; Fri–Sat to 1 AM. Kitchen Sun–Thu to 10 PM, Fri–Sat to 11 PM': 'Sun–Tue 11:30 AM–11 PM; Wed–Thu 11:30 AM–12 AM; Fri–Sat 11:30 AM–1 AM',
    'Mon–Sat seatings 6 PM & 8:30 PM': 'Mon–Sat 6–10:30 PM',
    'Mon & Wed–Sun 5–10 PM; closed Tue': 'Mon & Wed–Sun 5–10 PM',
    'Uber Eats lists 12–3:30 PM & 4:30 PM–12 AM daily; other listings say 11 AM–10 PM — confirm': 'Daily 12–3:30 PM & 4:30 PM–12 AM',
    'Daily 8–11:30 AM breakfast; 11:30 AM–10 PM lunch/dinner': 'Daily 8 AM–10 PM',
}

TIME = r'(?:noon|midnight|\d{1,2}(?::\d{2})?\s*(?:AM|PM)?)'
RANGE = re.compile(r'(' + TIME + r')\s*(?:–|-|to)\s*(' + TIME + r')')
DAYTOK = r'(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)'
DAYSPEC = re.compile(r'^\s*(Daily|' + DAYTOK + r'(?:\s*–\s*' + DAYTOK + r')?(?:\s*&\s*' + DAYTOK + r'(?:\s*–\s*' + DAYTOK + r')?)*)\b')


def tparse(tok):
    tok = tok.strip()
    if tok == 'noon':
        return 720, 'PM'
    if tok == 'midnight':
        return 0, 'AM'
    m = re.match(r'(\d{1,2})(?::(\d{2}))?\s*(AM|PM)?', tok)
    h, mi, mer = int(m.group(1)), int(m.group(2) or 0), m.group(3)
    return (h % 12) * 60 + mi, mer


def to24(mins, mer):
    return mins + (720 if mer == 'PM' else 0)


def parse_range(a, b):
    sm, smer = tparse(a)
    em, emer = tparse(b)
    if a.strip() in ('noon', 'midnight'):
        start = sm
    else:
        start = None
    end = to24(em, emer) if b.strip() not in ('noon', 'midnight') else em
    if start is None:
        if smer:
            start = to24(sm, smer)
        else:
            cand = to24(sm, emer)
            start = cand if cand <= end or end < 360 else to24(sm, 'AM')
            if emer == 'PM' and cand > end:
                start = to24(sm, 'AM')
    if end <= start:
        end += 1440
    return [start, end]


def days_of(spec):
    if spec == 'Daily':
        return list(range(7))
    out = []
    for part in spec.split('&'):
        part = part.strip()
        if '–' in part:
            a, b = [p.strip() for p in part.split('–')]
            i = DI[a]
            while True:
                out.append(i)
                if i == DI[b]:
                    break
                i = (i + 1) % 7
        else:
            out.append(DI[part])
    return out


def parse_hours(text):
    if not text:
        return None
    t = OVERRIDES.get(text, text)
    t = re.sub(r'\([^)]*\)', '', t)
    sched = [[] for _ in range(7)]
    last_days = None
    found = False
    for seg in re.split(r'[;.]', t):
        seg = seg.strip()
        if not seg or seg.lower().startswith('closed'):
            continue
        seg = re.sub(r'^(Café|Cafe|Lunch|Dinner|Kitchen)\s+', '', seg, flags=re.I)
        m = DAYSPEC.match(seg)
        if m:
            days = days_of(m.group(1))
            rest = seg[m.end():]
            last_days = days
        else:
            days = last_days
            rest = seg
        if days is None:
            continue
        for a, b in RANGE.findall(rest):
            iv = parse_range(a, b)
            for d in days:
                sched[d].append(iv)
            found = True
    if not found:
        return None
    # merge overlapping intervals
    for d in range(7):
        ivs = sorted(sched[d])
        merged = []
        for s, e in ivs:
            if merged and s <= merged[-1][1]:
                merged[-1][1] = max(merged[-1][1], e)
            else:
                merged.append([s, e])
        sched[d] = merged
    return sched


CITY_TZ = {
    'New York City': 'America/New_York', 'Chicago': 'America/Chicago', 'London': 'Europe/London',
    'Washington, DC': 'America/New_York', 'Aspen': 'America/Denver', 'Cleveland': 'America/New_York',
    'Durham': 'America/New_York', 'Las Vegas': 'America/Los_Angeles', 'San Francisco': 'America/Los_Angeles',
    'Denver': 'America/Denver', 'Seattle': 'America/Los_Angeles', 'College Park': 'America/New_York',
    'Miami': 'America/New_York', 'Paris': 'Europe/Paris', 'Malibu': 'America/Los_Angeles', 'Palm Springs': 'America/Los_Angeles', 'Newport Beach': 'America/Los_Angeles',
}

wb = openpyxl.load_workbook(SRC)
tabs = []
for ws in wb:
    hdr = [c.value for c in ws[1]]
    hcols = [i for i, h in enumerate(hdr) if h and 'Hours' in h]
    hcol = hcols[-1]
    col = {h: i for i, h in enumerate(hdr) if h}
    section = None
    rows = []
    title = None
    for r in range(2, ws.max_row + 1):
        cells = list(ws[r])
        v = [c.value for c in cells]
        if not any(v):
            continue
        name = v[0]
        if isinstance(name, str) and name.startswith('  ') and not v[1]:
            label = name.strip()
            if title is None:
                title = label
            if not label.startswith('Deep Canyon') and label not in ('Restaurants by City',):
                section = label
            continue
        if not v[1] or not name or str(name).startswith('©') or str(name).startswith('Delivery zones'):
            continue
        link = cells[len(hdr) - 1].hyperlink.target if cells[len(hdr) - 1].hyperlink else None
        hours = v[hcol]
        rec = {
            'name': str(name).strip(), 'cuisine': v[1], 'hood': v[2], 'address': v[3], 'phone': v[4],
            'hours': hours, 'sched': parse_hours(hours),
            'cls': v[col['Classification']], 'menu': link,
            'section': section, 'tz': CITY_TZ.get(section, 'America/Los_Angeles'),
            'confirm': bool(hours and re.search(r'confirm|vary|listing', hours, re.I)),
        }
        if 'Delivery Options' in col:
            rec['delivery'] = v[col['Delivery Options']]
            d = v[col['Approx. Distance']] or ''
            m = re.search(r'(\d+)', d)
            rec['dist'] = d
            rec['distN'] = int(m.group(1)) if m else 99
        rows.append(rec)
    tabs.append({'tab': ws.title, 'rows': rows})

json.dump(tabs, open(OUT, 'w'), ensure_ascii=False)

# report
bad = []
for t in tabs:
    for r in t['rows']:
        if r['hours'] and not r['sched']:
            bad.append(r['hours'])
print('tabs', [(t['tab'], len(t['rows'])) for t in tabs])
print('unparsed', bad)
