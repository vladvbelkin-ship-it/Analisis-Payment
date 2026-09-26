# -*- coding: utf-8 -*-
"""Сопоставление входящих платежей из PDF с базами клиентов (Вольная / Удальцова).

Ключи (по убыванию надёжности):
 K1 телефон отправителя (СБП в Альфа) == Телефон клиента                     -> 100%
 K3 ФИО: совпали фамилия ИЛИ имя+отчество отправителя/назначения             -> 90%
    (+совпало и то и другое -> 95%)
 K4 только имя (без фамилии), часто + отчество                               -> 60%
 K5 только сумма, краткая тарифу                                             -> не относим к центру (25%), в TSV не пишем
Доп. подтверждения дописываются в примечание: тариф, "есть и в др. центре".
"""
import json, re, sys, datetime
from collections import defaultdict
import openpyxl

CENTERS = {'Вольная': 'Учет Вольная.xlsx', 'Удальцова': 'Учет Удальцова.xlsx'}
BANKS = [('АльфаБанк р/с ...7556', 'out/alfa_incoming.json'),
         ('Сбербанк карта ...9110', 'out/sber_9110_incoming.json'),
         ('Сбербанк карта ...4031', 'out/sber_4031_incoming.json')]

TARIFF_SINGLE = 2900
TARIFF_DISC = 2700            # занятия по абонементу 8 (21600) или 10 (27000)
PRICE_SETS = {21600: 'абонемент 8×2700', 27000: 'абонемент 10×2700',
              19440: 'абонемент Вольная', 17280: 'абонемент Вольная'}


def norm(s):
    if not s:
        return ''
    s = str(s).lower().replace('ё', 'е')
    return ''.join(c for c in s if c.isalpha())


def tel(v):
    if v is None:
        return ''
    if isinstance(v, float):
        v = str(int(v))
    d = re.sub(r'\D', '', str(v))
    if len(d) == 11 and d[0] == '8':
        d = '7' + d[1:]
    if len(d) == 10:
        d = '7' + d
    return d if len(d) == 11 else ''


STOP = {'занятие', 'занятия', 'занятий', 'абонемент', 'оплата', 'перевод', 'спасибо',
        'месяц', 'неделя', 'две', 'три', 'четыре', 'суббота', 'воскресенье', 'среда',
        'четверг', 'вторник', 'пятница', 'понедельник', 'подарок', 'билет', 'карт',
        'счет', 'руб', 'россия', 'москва', 'центр', 'учитель', 'педагог', 'минус',
        'плюс', 'итого', 'всего', 'дата', 'врем', 'бннс', 'ндс', 'без', 'пн', 'вт',
        'ср', 'чт', 'пт', 'сб', 'вс', 'для', 'от', 'за', 'на', 'из', 'по', 'и', 'с',
        'в', 'нум', 'nomer', 'schet'}
MAINTOKENS = {'александр', 'александра', 'екатерина', 'мария', 'анна', 'елена', 'ольга',
              'наталья', 'сергей', 'андрей', 'дмитрий', 'maksim', 'максим', 'иван',
              'татьяна', 'юлия', 'иррина', 'ирина', 'светлана', 'анастасия'}

ENDINGS = ['ами', 'ями', 'овна', 'евна', 'ична', 'ична', 'ых', 'их', 'ой', 'ей', 'ова',
           'ева', 'ов', 'ев', 'ам', 'ям', 'ах', 'ях', 'ую', 'юю', 'ого', 'его', 'ому',
           'ему', 'ыми', 'ими', 'ах', 'у', 'ю', 'ы', 'и', 'а', 'я', 'о', 'е']


def stems(tok):
    res = {tok}
    for e in ENDINGS:
        if tok.endswith(e) and len(tok) - len(e) >= 4:
            res.add(tok[:-len(e)])
    # специфично для русских фамилий/отчеств: removal of longer endings first handled by set
    return res


def words(s):
    out = []
    for w in re.split(r'[^\w]+', str(s or '')):
        n = norm(w)
        if len(n) >= 4 and n not in STOP:
            out.append(n)
    return out


# ---------------- загрузка клиентов ----------------
clients = {}
for center, fn in CENTERS.items():
    wb = openpyxl.load_workbook(fn, data_only=True)
    ws = wb['Клиенты']
    hdr = None
    seen = {}
    for r in ws.iter_rows(values_only=True):
        if r and isinstance(r[0], str) and r[0].strip() == 'Имя':
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
        key = (nm.strip(), par, t, req)
        if key in seen:
            continue
        # разбор имён на токены: ребёнок и родитель
        ch_w = [w for w in words(nm)]
        pa_w = [w for w in words(par)]
        seen[key] = {'child': nm.strip(), 'parent': par, 'tel': t, 'req': req,
                     'raz': d.get('Стоимость разового'), 'ab': d.get('Стоимость абонемента'),
                     'ch_words': ch_w, 'pa_words': pa_w}
    clients[center] = list(seen.values())

TEL_IDX = defaultdict(set)
NAME_IDX = defaultdict(set)      # stem -> {(center, child, parent)}
for center, rows in clients.items():
    for c in rows:
        ident = (center, c['child'], c['parent'])
        if c['tel']:
            TEL_IDX[c['tel']].add(ident)
        for tk in set(c['ch_words'] + c['pa_words']):
            for s in stems(tk):
                NAME_IDX[s].add(ident)

SURNAME_HINT = re.compile(r'[а-яё]+$')  # все слова — кириллица; используем длину/позицию


def classify_tokens(text_words):
    """Возвращает hits: ident -> {'surname':bool,'firstname':bool,'patronymic':bool,'tokens':set}"""
    hits = defaultdict(lambda: {'sur': set(), 'name': set(), 'patr': set()})
    for pos, tk in enumerate(text_words):
        for s in stems(tk):
            for ident in NAME_IDX.get(s, ()):
                # определим роль токена: есть ли он в имени ребёнка/родителя как первое слово (=имя)
                # и как второе/последнее (=фамилия)
                role = None
                ch, pa = ident[1], ident[2]
                chw = words(ch); paw = words(pa)
                if tk in chw[1:] or tk in paw[1:]:
                    role = 'sur'
                elif tk in chw[:1] or tk in paw[:1]:
                    role = 'name'
                elif tk in chw[2:] or tk in paw[2:]:
                    role = 'patr'
                else:
                    role = 'sur' if len(tk) >= 5 else 'name'
                hits[ident][role].add(tk)
    return hits


def price_note(amount):
    a = round(amount)
    if a in PRICE_SETS:
        return PRICE_SETS[a]
    if a % TARIFF_SINGLE == 0:
        return f'{a // TARIFF_SINGLE}×{TARIFF_SINGLE} (разовые)'
    if a % TARIFF_DISC == 0:
        return f'{a // TARIFF_DISC}×{TARIFF_DISC} (абонемент)'
    return ''


def pdate(d):
    dd, mm, yy = d.split('.')
    return datetime.date(int(yy), int(mm), int(dd))


# ---- уже учтённые оплаты (для пометки) ----
existing = defaultdict(list)
for center, fn in CENTERS.items():
    wb = openpyxl.load_workbook(fn, data_only=True)
    ws = wb['Оплаты']
    hdr = None
    for r in ws.iter_rows(values_only=True):
        if r and isinstance(r[0], str) and r[0].strip() == 'Дата':
            hdr = list(r)
            continue
        if hdr is None or not r or r[hdr.index('Сумма')] is None:
            continue
        d = dict(zip(hdr, r))
        dt = d.get('Дата')
        ds = dt.strftime('%d.%m.%Y') if hasattr(dt, 'strftime') else str(dt)[:10]
        try:
            existing[center].append((pdate(ds), str(d.get('Имя ребенка') or '').strip(),
                                     round(float(d['Сумма']))))
        except Exception:
            pass

RESULT = []
for bank, path in BANKS:
    for o in json.load(open(path, encoding='utf-8')):
        amount = o['amount']
        if amount <= 0:
            continue
        crit, prob, idents = [], 0, set()
        phone = o.get('phone') or ''
        if phone and phone in TEL_IDX:
            idents |= TEL_IDX[phone]
            crit.append(f'K1 телефон {phone}')
            prob = 100

        sender = o.get('sender') or ''
        purpose = o.get('purpose') or ''
        via = o.get('via_bank') or ''
        tw = words(sender) + words(purpose)
        if tw:
            hits = classify_tokens(tw)
            strong = {i: h for i, h in hits.items() if h['sur'] or (h['name'] and h['patr'])}
            weak = {i: h for i, h in hits.items() if i not in strong}
            if strong:
                idents |= set(strong)
                best = max(strong, key=lambda k: (len(strong[k]['sur']), len(strong[k]['name'])))
                hs, hn = strong[best]['sur'], strong[best]['name']
                lbl = '/'.join(sorted(hs | hn | strong[best]['patr']))
                if hs and hn:
                    prob = max(prob, 95)
                    crit.append(f'K3 ФИО+фамилия "{best[1]}" ({lbl})')
                elif hs:
                    prob = max(prob, 90)
                    crit.append(f'K3 фамилия "{best[1]}" ({lbl})')
                else:
                    prob = max(prob, 90)
                    crit.append(f'K3 имя+отчество "{best[1]}" ({lbl})')
            elif weak and prob < 90:
                # слабое совпадение: принимаем только если это распространённое имя и сумма тарифная
                pn = price_note(amount)
                usable = {i: h for i, h in weak.items() if h['name']}
                if usable and pn:
                    idents |= set(usable)
                    best = max(usable, key=lambda k: len(usable[k]['name']))
                    crit.append(f'K4 только имя "{best[1]}" ({" ".join(sorted(usable[best]["name"]))})')
                    prob = max(prob, 60)

        centers = sorted({i[0] for i in idents})
        children = sorted({i[1] for i in idents})
        parents = sorted({i[2] for i in idents if i[2]})
        pn = price_note(amount)
        if not idents and pn:
            crit.append(f'K5 только сумма ({pn}) — центр не определён')
            prob = 25

        already = []
        for c in centers:
            for (dd, ch, sa) in existing[c]:
                if sa == round(amount) and abs((dd - pdate(o['date'])).days) <= 3:
                    already.append(f'уже учтено в {c}: {dd:%d.%m.%Y} {ch}')
                    break

        RESULT.append({'bank': bank, 'date': o['date'], 'amount': amount,
                       'phone': phone, 'sender': sender, 'purpose': purpose, 'via': via,
                       'centers': centers, 'children': children, 'parents': parents,
                       'crit': '; '.join(crit), 'prob': prob, 'tariff': pn,
                       'already': already[:2]})

json.dump(RESULT, open('out/payments_matched2.json', 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)

matched = [r for r in RESULT if r['centers']]
print('всего входящих:', len(RESULT), '| сопоставлено с клиентами:', len(matched))
byc = defaultdict(int)
for r in matched:
    for c in r['centers']:
        byc[c] += 1
print('по центрам:', dict(byc))
pd = defaultdict(int)
for r in RESULT:
    pd[r['prob']] += 1
print('по вероятности:', dict(sorted(pd.items())))
print('\n=== примеры prob>=90 ===')
for r in [x for x in RESULT if x['prob'] >= 90][:25]:
    print(r['bank'], r['date'], r['amount'], '|', r['centers'], '|', r['crit'][:120])
print('\n=== примеры prob==60 ===')
for r in [x for x in RESULT if x['prob'] == 60][:20]:
    print(r['bank'], r['date'], r['amount'], '|', r['centers'], '|', r['crit'][:120])
