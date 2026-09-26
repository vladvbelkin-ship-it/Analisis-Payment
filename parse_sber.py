# -*- coding: utf-8 -*-
"""Парсинг выписок Сбер (карты 9110 и 4031): входящие переводы -> out/sber_<card>_incoming.json"""
import re, json, pdfplumber

FILES = {'Выписка по счёту дебетовой карты (1).pdf': '9110',
         'Выписка по счёту дебетовой карты (3).pdf': '4031'}

def num(s):
    return float(s.replace('\u00a0','').replace(' ','').replace(',','.'))

AMT = r'([\d\s\u00a0]*\d,\d{2})'
HEAD = re.compile(rf'^(\d{{2}}\.\d{{2}}\.\d{{4}})\s+(\d{{2}}:\d{{2}})\s+(.+?)\s+\+?{AMT}\s+{AMT}\s*$')
CONT = re.compile(r'^(\d{2}\.\d{2}\.\d{4})\s+(\d{6})\s+(.*)$')

for fn, card in FILES.items():
    with pdfplumber.open(fn) as pdf:
        lines = []
        for p in pdf.pages:
            lines += (p.extract_text() or '').splitlines()
    ops, cur = [], None
    skip = re.compile(r'(Страница \d+ из|Продолжение на следующей|Для проверки подлинности|Зайдите в приложение|Нажмите кнопку|документ в электронном|Действителен|до \d{2}\.\d{2}|Предоставляя QR|Расшифровка операций|ДАТА ОПЕРАЦИИ|Дата обработки|и код авторизации|Остаток на|Номер счёта|Валюта Российский|Дата открытия|Дата закрытия|Карта |За период|ИТОГО ПО|Владелец счёта|www.sberbank.ru|ул. Вавилова|Выписка по счёту|^900 )')
    for raw in lines:
        l = raw.strip()
        if not l or skip.search(l): continue
        m = HEAD.match(l)
        if m:
            if cur: ops.append(cur)
            cur = {'date': m.group(1), 'time': m.group(2), 'cat': m.group(3),
                   'amount': num(m.group(4)), 'balance': num(m.group(5)), 'desc': ''}
            continue
        m = CONT.match(l)
        if m and cur is not None:
            cur['proc'] = m.group(1); cur['desc'] += ' ' + m.group(3)
        elif cur is not None:
            cur['desc'] += ' ' + l
    if cur: ops.append(cur)

    inc = []
    for o in ops:
        d = o.get('desc','').strip()
        if not re.search(r'Перевод (от|из)\b', d):   # «Перевод для/в» — исходящие
            continue
        ms = re.search(r'Перевод от ([А-ЯЁ])\.\s*([А-ЯЁ][а-яё]+)\s*([А-ЯЁ][а-яё]+)?', d)
        sender = ((ms.group(1)+'. '+ms.group(2)+' '+(ms.group(3) or '')).strip()) if ms else ''
        mv = re.search(r'Перевод из ([^.]+?)\.', d)
        via = mv.group(1).strip() if mv else ''
        ttype = 'card' if 'Перевод на карту' in o['cat'] else ('sbp' if 'Перевод СБП' in o['cat'] else 'other')
        inc.append({'date': o['date'], 'time': o['time'], 'type': ttype,
                    'amount': o['amount'], 'sender': sender, 'via_bank': via,
                    'balance': o['balance'], 'desc': d})
    json.dump(inc, open(f'out/sber_{card}_incoming.json','w',encoding='utf-8'), ensure_ascii=False, indent=1)
    big = [x for x in inc if round(x['amount']) >= 2500]
    print(card, '| строк операций:', len(ops), '| входящих:', len(inc), '| >=2500:', len(big),
          '| с ФИО:', sum(1 for x in inc if x['sender']))
