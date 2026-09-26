# -*- coding: utf-8 -*-
"""Сопоставление входящих платежей из PDF-выписок с базами клиентов (Вольная / Удальцова).

Ключи сопоставления (по убыванию надёжности):
  K1 телефон отправителя (СБП в Альфа) == Телефон клиента                     -> 100%
  K2 последние 4 цифры карты плательщика (Реквизиты плательщика) в тексте     -> 95%
  K3 ФИО ребёнка/родителя в назначении платежа или в имени отправителя        -> 85% (фамилия+имя), 70% (только фамилия)
  K4 только имя (без фамилии) + совпадение по тарифу                          -> 50%
  K5 только сумма, краткая тарифу (2900*N разовое; 2700*N абонемент: 8->21600, 10->27000) -> 30%, не основание
"""
import json, re, sys
from collections import defaultdict
import openpyxl

CENTERS = {'Вольная': 'Учет Вольная.xlsx', 'Удальцова': 'Учет Удальцова.xlsx'}
BANKS = [('АльфаБанк р/с ...7556', 'out/alfa_incoming.json'),
         ('Сбербанк карта 9110',   'out/sber_9110_incoming.json'),
         ('Сбербанк карта 4031',   'out/sber_4031_incoming.json')]

TARIFF_SINGLE = 2900
TARIFF_LESSON_DISCOUNT = 2700          # абонемент 8 (21600) или 10 (27000) занятий
PRICE_SETS = {21600: '8×2700', 27000: '10×2700', 19440: 'абонемент Вольная',
              17280: 'абонемент Вольная', 2900: 'разовое', 5800: '2×2900', 8700: '3×2900',
              11600: '4×2900', 14500: '5×2900', 17400: '6×2900', 20300: '7×2900',
              8100: '3×2700', 10800: '4×2700', 13500: '5×2700', 16200: '6×2700', 18900: '7×2700'}


def norm(s):
    if not s:
        return ''
    s = str(s).lower().replace('ё', 'е')
    return ''.join(c for c in s if c.isalpha())


def tel(v):
    if v is None:
        return ''
    d = re.sub(r'\D', '', str(v))
    if len(d) == 11 and d[0] == '8':
        d = '7' + d[1:]
    if len(d) == 10:
        d = '7' + d
    return d if len(d) == 11 else ''


STOP = {'занятия', 'занятий', 'абонемент', 'оплата', 'перевод', 'за', 'за ', 'спасибо',
        'безндс', 'ндс', 'месяц', 'неделя', 'две', 'три', 'четыре', 'суббота', 'среда',
        'четверг', 'вторник', 'пятница', 'понедельник', 'подарок', 'на', 'от', 'для'}


def tokens(s):
    out = []
    for w in re.split(r'[^\w]+', str(s or '')):
        n = norm(w)
        if len(n) >= 4 and n not in STOP:
            out.append(n)
    return out


# ---------- загрузка клиентов ----------
clients = {}
for center, fn in CENTERS.items():
    wb = openpyxl.load_workbook(fn, data_only=True)
    ws = wb['Клиенты']
    hdr = None
    seen = {}
    for r in ws.iter_rows(values_only=True):
        if r and r[0] == 'Имя':
            hdr = list(r)
            continue
        if hdr is None:
            continue
        d = dict(zip(hdr, r))
        nm = d.get('Имя')
        if not isinstance(nm, str) or not nm.strip():
            continue
        par = str(d.get('Имя родителя') or '').strip()
        t = tel(d.get('Телефон'))
        req = str(d.get('Реквизиты плательщика') or '').strip()
        raz = d.get('Стоимость разового')
        ab = d.get('Стоимость абонемента')
        key = (nm.strip(), par, t, req)
        if key in seen:
            if raz and not seen[key]['raz']:
                seen[key]['raz'] = raz
            if ab and not seen[key]['ab']:
                seen[key]['ab'] = ab
            continue
        seen[key] = {'child': nm.strip(), 'parent': par, 'tel': t, 'req': req,
                     'raz': raz, 'ab': ab}
    clients[center] = list(seen.values())

TEL_IDX = defaultdict(set)      # phone -> {(center, child, parent)}
CARD_IDX = defaultdict(set)     # last4 -> {(center, child, parent)}
NAME_IDX = defaultdict(set)     # token -> {(center, child, parent)}
for center, rows in clients.items():
    for c in rows:
        ident = (center, c['child'], c['parent'])
        if c['tel']:
            TEL_IDX[c['tel']].add(ident)
        m = re.fullmatch(r'\d{4}', c['req'])
        if m:
            CARD_IDX[m.group(0)].add(ident)
        for src in (c['child'], c['parent']):
            for tk in tokens(src):
                NAME_IDX[tk].add(ident)

ENDINGS = ['ами', 'ями', 'овна', 'евна', 'ична', 'ых', 'их', 'ой', 'ей', 'ов', 'ев', 'ева',
           'ова', 'ам', 'ям', 'ах', 'ях', 'ах', 'у', 'ю', 'ы', 'и', 'а', 'я', 'о', 'е']


def stems(tok):
    """основы слова для матчинга без учёта падежа"""
    res = {tok}
    for e in ENDINGS:
        if e and tok.endswith(e) and len(tok) - len(e) >= 4:
            res.add(tok[:-len(e)])
    return res


IDX_STEM = defaultdict(set)
for tk, idents in NAME_IDX.items():
    for s in stems(tk):
        IDX_STEM[s] |= idents


def lookup(text):
    """возвращает (idents, n_matched_tokens, matched_tokens) по всем основам токенов текста"""
    hits = defaultdict(int)
    detail = defaultdict(set)
    for tk in tokens(text):
        found = set()
        for s in stems(tk):
            if s in IDX_STEM:
                found |= IDX_STEM[s]
        for f in found:
            hits[f] += 1
            detail[f].add(tk)
    return hits, detail


def price_note(amount):
    a = round(amount)
    if a in PRICE_SETS:
        return PRICE_SETS[a]
    if a % TARIFF_SINGLE == 0:
        return f'{a // TARIFF_SINGLE}×2900'
    if a % TARIFF_LESSON_DISCOUNT == 0:
        return f'{a // TARIFF_LESSON_DISCOUNT}×2700'
    return ''


def parse_date(d):
    dd, mm, yy = d.split('.')
    return int(yy) * 10000 + int(mm) * 100 + int(dd)


# ---------- уже внесённые оплаты (для пометки "уже учтено") ----------
existing = defaultdict(list)     # (center) -> [(date, child, amount)]
for center, fn in CENTERS.items():
    wb = openpyxl.load_workbook(fn, data_only=True)
    ws = wb['Оплаты']
    hdr = None
    for r in ws.iter_rows(values_only=True):
        if r and r[0] == 'Дата':
            hdr = list(r)
            continue
        if hdr is None:
            continue
        d = dict(zip(hdr, r))
        if d.get('Дата') and d.get('Сумма'):
            dt = d['Дата']
            ds = dt.strftime('%d.%m.%Y') if hasattr(dt, 'strftime') else str(dt)
            existing[center].append((ds, str(d.get('Имя ребенка') or '').strip(), round(float(d['Сумма']))))

RESULT = []
for bank, path in BANKS:
    for o in json.load(open(path, encoding='utf-8')):
        amount = o['amount']
        if amount <= 0:
            continue
        text_bits = []
        crit = []
        prob = 0
        idents = set()

        phone = o.get('phone') or ''
        if phone and phone in TEL_IDX:
            idents |= TEL_IDX[phone]
            crit.append(f'K1 телефон {phone}')
            prob = 100
        # последние 4 цифры карты плательщика из текста выписки
        desc = o.get('desc') or ''
        fours = set(re.findall(r'(?:карты|карт)\s*[«"„]?\D{0,3}(\d{4})', desc))
        fours |= set(re.findall(r'\*{2,6}\s*(\d{4})', desc))
        card_hits = {f: CARD_IDX[f] for f in fours if f in CARD_IDX}
        if card_hits:
            for f, ids in card_hits.items():
                idents |= ids
                crit.append(f'K2 карта ...{f}')
            prob = max(prob, 95)

        search_text = ' '.join([o.get('purpose') or '', o.get('sender') or '', o.get('via_bank') or ''])
        if prob < 95 and search_text.strip():
            hits, detail = lookup(search_text)
            strong = {i: n for i, n in hits.items() if n >= 2}
            weak = {i: n for i, n in hits.items() if n == 1}
            if strong:
                idents |= set(strong)
                best = max(strong, key=lambda k: strong[k])
                crit.append(f'K3 ФИО "{best[1]} {best[2]}" ({"/".join(sorted(detail[best]))})')
                prob = max(prob, 85)
            elif weak and price_note(amount):
                idents |= set(weak)
                best = max(weak, key=lambda k: weak[k])
                crit.append(f'K3 частичное ФИО "{best[1]} {best[2]}" ({"/".join(sorted(detail[best]))}) + тариф')
                prob = max(prob, 60)
            elif weak:
                idents |= set(weak)
                best = max(weak, key=lambda k: weak[k])
                crit.append(f'K4 только имя "{best[1]} {best[2]}" ({"/".join(sorted(detail[best]))})')
                prob = max(prob, 40)

        if not idents and price_note(amount):
            crit.append(f'K5 только сумма, тариф {price_note(amount)}')
            prob = 25

        centers = sorted({i[0] for i in idents})
        children = sorted({f"{i[1]}" + (f" (род. {i[2]})" if i[2] else '') for i in idents})
        pn = price_note(amount)

        already = []
        for c in centers:
            for (ds, ch, sa) in existing[c]:
                if abs(sa - round(amount)) < 1 and parse_date(ds) and abs(parse_date(ds) - parse_date(o['date'])) <= 3:
                    if not children or any(ch.lower().split()[0][:4] in ' '.join(children).lower() for ch in [ch] if ch):
                        already.append(f'{c}:{ds} {ch} {sa}')
                        break

        RESULT.append({'bank': bank, 'date': o['date'], 'time': o.get('time', ''),
                       'amount': amount, 'phone': phone, 'sender': o.get('sender', ''),
                       'purpose': o.get('purpose', ''), 'centers': centers,
                       'children': children, 'crit': '; '.join(crit), 'prob': prob,
                       'tariff': pn, 'already': already[:3], 'desc': desc[:120]})

json.dump(RESULT, open('out/payments_matched.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

matched = [r for r in RESULT if r['centers']]
print('всего входящих операций:', len(RESULT))
print('сопоставленных с клиентами:', len(matched))
byc = defaultdict(int)
for r in matched:
    for c in r['centers']:
        byc[c] += 1
print('по центрам:', dict(byc))
print('по вероятности:', dict(sorted(defaultdict(lambda: 0, {p: sum(1 for r in RESULT if r['prob'] == p)
                                                              for p in {r['prob'] for r in RESULT}}).items())))
print('не сопоставлено:', len(RESULT) - len(matched))
