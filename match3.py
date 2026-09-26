# -*- coding: utf-8 -*-
"""Финальное сопоставление входящих платежей из PDF с базами клиентов.

Ключи надёжности (вероятность в %):
 K1 телефон отправителя (СБП, Альфа) == Телефон клиента .................. 100%
 K2 последние 4 цифры карты плательщика ("Реквизиты плательщика") ......... 95%
    (в текущих выписках номера карт плательщиков не печатаются — ключ готов,
     срабатывает только если 4-значный код из «Реквизитов» встречается в тексте
     строки платежа; сам текст карты получателя (****9110/****4031/р/с ...7556)
     исключён из сравнения)
 K3 совпадение ФИО отправителя / назначения платежа с именем ребёнка или
    родителя (фамилия ИЛИ имя+отчество, с учётом падежей) ................ 90%
    + подтверждение суммой (тариф) ....................................... 95%
 K4 единственное слабое совпадение (только одно имя, без фамилии),
    если кандидат ровно один и сумма тарифная ........................... 60%
 K5 только сумма, краткая тарифу (2900×N разовые; 2700×N абонемент:
    8 занятий = 21600, 10 занятий = 27000) — центр НЕ определяется ...... 25%
    (в TSV по центрам не включается)
Если клиент найден в обеих базах — строка дублируется в оба TSV с пометкой
"есть и в др. центре".
"""
import json, re, datetime
from collections import defaultdict
import openpyxl

CENTERS = {'Вольная': 'Учет Вольная.xlsx', 'Удальцова': 'Учет Удальцова.xlsx'}
BANKS = [('АльфаБанк р/с ...7556', 'out/alfa_incoming.json'),
         ('Сбербанк карта ...9110', 'out/sber_9110_incoming.json'),
         ('Сбербанк карта ...4031', 'out/sber_4031_incoming.json')]

TARIFF_SINGLE = 2900          # одно занятие без абонемента
TARIFF_DISC = 2700            # занятие по абонементу 8 (21600) или 10 (27000)
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
        'счет', 'руб', 'россия', 'москва', 'центр', 'учитель', 'педагог', 'итого',
        'всего', 'дата', 'ндс', 'без', 'для', 'от', 'за', 'на', 'из', 'по', 'нум'}

ENDINGS = ['овна', 'евна', 'ична', 'ично', 'ины', 'ые', 'ога', 'ега', 'огу', 'егу',
           'ом', 'ем', 'ов', 'ев', 'ам', 'ям', 'ах', 'ях', 'ой', 'ей', 'ий', 'ый',
           'ая', 'яя', 'ое', 'ее', 'ых', 'их', 'ым', 'им', 'у', 'ю', 'ы', 'и', 'а',
           'я', 'о', 'е']


def stems(tok):
    res = {tok}
    for e in sorted(ENDINGS, key=len, reverse=True):
        if tok.endswith(e) and len(tok) - len(e) >= 4:
            res.add(tok[:-len(e)])
    return res


def words(s):
    out = []
    for w in re.split(r'[^\w]+', str(s or '')):
        n = norm(w)
        if len(n) >= 4 and n not in STOP:
            out.append(n)
    return out


# ---------------- клиенты ----------------
clients = {}
for center, fn in CENTERS.items():
    wb = openpyxl.load_workbook(fn, data_only=True)
    ws = wb['Клиенты']
    hdr, seen = None, {}
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
        chw, paw = words(nm), words(par)
        seen[key] = {'child': nm.strip(), 'parent': par, 'tel': t, 'req': req,
                     'ch_set': set(chw), 'pa_set': set(paw)}
    clients[center] = list(seen.values())

PATR_SET = {'ивановна','петровна','кузьмина'}
NAMES = set("""александр александра алексей алёна анастасия андрей Анна антон артем
арсений борис валерия вадим владимир Владислав виталий виктор галина дарья дмитрий
евгений евгения егор екатерина елена иван игорь Ильяс Инна Кирилл Кристина кузьма
лукерья марат маргарита марк максим михаил нара никита нина оксана олег павел
поликарп роман рукамина светлана сергей степан татьяна тимур ульяна федор хабугович
эдуард юлий ярослав""".lower().split())
PATR_PREFIX = ('роман', 'серг', 'андр', 'владимир', 'алекс', 'николай', 'геннад',
               'михайл', 'дмитр', 'игор', 'павл', 'олегов', 'виктор', 'борис',
               'руслан', 'эдуард', 'константин', 'герман', 'валер', 'влас', 'клим')

TEL_IDX = defaultdict(set)
CARD_IDX = defaultdict(set)
NAME_IDX = defaultdict(set)   # stem -> idents
ROLE = {}                     # ident -> set of all name-stems
SUR_IDX = defaultdict(set)    # ident -> фамильные стемы
FIRST_IDX = defaultdict(set)  # ident -> стемы имён
PATR_IDX = defaultdict(set)   # ident -> стемы отчеств
for center, rows in clients.items():
    for c in rows:
        ident = (center, c['child'], c['parent'])
        if c['tel']:
            TEL_IDX[c['tel']].add(ident)
        if re.fullmatch(r'\d{4}', c['req']):
            CARD_IDX[c['req']].add(ident)
        allstems = set()
        sur_stems = set()
        for tk in (c['ch_set'] | c['pa_set']):
            for s in stems(tk):
                NAME_IDX[s].add(ident)
                allstems.add(s)
        # роль токена: отчество / имя / фамилия
        for tk in (c['ch_set'] | c['pa_set']):
            tks = stems(tk)
            low = tk.lower()
            is_patr = (low.endswith(('ич', 'ит', 'вна', 'чна')) and any(
                low.startswith(p) for p in PATR_PREFIX)) or low in PATR_SET
            is_first = low in NAMES
            if is_patr:
                PATR_IDX[ident] |= tks
            elif is_first:
                FIRST_IDX[ident] |= tks
            else:
                SUR_IDX[ident] |= tks
                sur_stems |= tks
        ROLE[ident] = allstems




def is_name(tok):
    return tok in {n.lower() for n in NAMES} or any(
        tok.startswith(p) and (tok.endswith(('ич', 'ит', 'ична')) or len(tok) < 9)
        for p in ())   # placeholder


def match_names(text_words, sender_first=None):
    """Сопоставление с учётом ролей. Возвращает ident -> число значимых совпадений.

    Значимое совпадение: фамилия (общая для ребёнка и родителя), имя или отчество.
    Токены из имени плательщика сверяются с именем/отчеством РОДИТЕЛЯ, токены из
    назначения платежа — с обоими именами.
    """
    hits = defaultdict(int)
    detail = defaultdict(set)
    twset = set(text_words)
    stems_of_text = {}
    for tk in text_words:
        stems_of_text[tk] = stems(tk)
    allstems = set()
    for v in stems_of_text.values():
        allstems |= v
    for tk in text_words:
        pass
    for ident in ROLE:
        center, child, parent = ident
        score = 0
        matched_tokens = set()
        for tk in text_words:
            st = stems_of_text[tk] & ROLE[ident]
            if not st:
                continue
            # роль токена в базе
            sur = st & SUR_IDX.get(ident, set())
            nm = st & FIRST_IDX.get(ident, set())
            pt = st & PATR_IDX.get(ident, set())
            if sur:
                score += 2
                matched_tokens.add(tk)
            elif pt:
                score += 1
                matched_tokens.add(tk)
            elif nm:
                # имя ребёнка часто совпадает с именем родителя-тёзки; считаем 1
                score += 1
                matched_tokens.add(tk)
        if score >= 2:
            hits[ident] = min(score, 4)
            detail[ident] = matched_tokens
    return hits, detail


def price_note(amount):
    a = round(amount)
    if a in PRICE_SETS:
        return PRICE_SETS[a]
    if a % TARIFF_SINGLE == 0:
        return f'{a // TARIFF_SINGLE}×{TARIFF_SINGLE} разов.'
    if a % TARIFF_DISC == 0:
        return f'{a // TARIFF_DISC}×{TARIFF_DISC} абан.'
    return ''


def pdate(d):
    dd, mm, yy = d.split('.')
    return datetime.date(int(yy), int(mm), int(dd))


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
        try:
            dd = dt.date() if hasattr(dt, 'strftime') else \
                datetime.datetime.strptime(str(dt)[:10], '%Y-%m-%d').date()
            existing[center].append((dd, str(d.get('Имя ребенка') or '').strip(),
                                     round(float(d['Сумма']))))
        except Exception:
            pass

RESULT = []
for bank, path in BANKS:
    own_cards = set(re.findall(r'\d{4}', bank))
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

        desc = (o.get('desc') or '')
        # K2: последние 4 цифры КАРТЫ ПЛАТЕЛЬЩИКА. В текущих выписках карты
        # плательщиков не печатаются; собственные реквизиты получателя
        # (р/с ...7556, ****9110, ****4031) из сравнения исключены.
        own_digits = set(own_cards)
        fours = set()
        for m in re.finditer(r'(?:карт(?:ы|ы плательщика|ой)\s*[«"„]?\D{0,3})(\d{4})', desc):
            if m.group(1) not in own_digits:
                fours.add(m.group(1))
        for m in re.finditer(r'\*{2,}\s*(\d{4})', desc):
            if m.group(1) not in own_digits:
                fours.add(m.group(1))
        for f in fours & set(CARD_IDX):
            idents |= CARD_IDX[f]
            crit.append(f'K2 посл.4 карты {f}')
            prob = max(prob, 95)

        tw = words(o.get('sender') or '') + words(o.get('purpose') or '')
        pn = price_note(amount)
        named = None                      # лучшая пара (ребёнок, родитель) по ФИО
        if tw:
            hits, det = match_names(tw)
            strong = {i: h for i, h in hits.items() if h >= 2}
            single = {i: h for i, h in hits.items() if h == 1}
            if strong:
                idents |= set(strong)
                best = max(strong, key=lambda k: strong[k])
                prob = max(prob, 95 if pn else 90)
                crit.append(f'K3 ФИО "{best[1]}"'
                            + (f' + сумма ({pn})' if pn else ''))
                named = (best[1], best[2])
            elif len(single) == 1 and pn:
                i = next(iter(single))
                idents.add(i)
                prob = max(prob, 60)
                crit.append(f'K4 единств. имя "{i[1]}" + тариф {pn}')
                named = (i[1], i[2])
            elif single:
                # несколько кандидатов по одному лишь имени — не относим
                pass

        # если телефон/карта дали несколько клиентов-тёзок, уточняем по тексту
        if len(idents) > 1 and named:
            pref = {i for i in idents if (i[1], i[2]) == named}
            if pref:
                idents = pref

        centers = sorted({i[0] for i in idents})
        children = sorted({i[1] for i in idents})
        parents = sorted({i[2] for i in idents if i[2]})
        if not idents and pn:
            crit.append(f'K5 только сумма ({pn}), центр не определён')
            prob = 25

        already = []
        for c in centers:
            for (dd, ch, sa) in existing[c]:
                if sa == round(amount) and abs((dd - pdate(o['date'])).days) <= 3:
                    already.append(f'уже учтено в {c}: {dd:%d.%m.%Y} {ch}')
                    break

        RESULT.append({'bank': bank, 'date': o['date'], 'amount': amount,
                       'phone': phone, 'sender': o.get('sender', ''),
                       'purpose': o.get('purpose', ''), 'via': o.get('via_bank', ''),
                       'centers': centers, 'children': children, 'parents': parents,
                       'crit': '; '.join(crit), 'prob': prob, 'tariff': pn,
                       'already': already[:2]})

json.dump(RESULT, open('out/payments_matched3.json', 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)

matched = [r for r in RESULT if r['centers']]
print('всего входящих:', len(RESULT), '| сопоставлено:', len(matched))
byc = defaultdict(int)
for r in matched:
    for c in r['centers']:
        byc[c] += 1
print('по центрам:', dict(byc))
pd = defaultdict(int)
for r in RESULT:
    pd[r['prob']] += 1
print('по вероятности:', dict(sorted(pd.items())))
