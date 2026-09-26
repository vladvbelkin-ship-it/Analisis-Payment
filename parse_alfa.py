# -*- coding: utf-8 -*-
"""Парсинг выписки Альфа-Банка: входящие платежи -> out/alfa_incoming.json"""
import re, json

RAW = open('pdf_alfa_raw.txt', encoding='utf-8').read()
DATE_RE = re.compile(r'^(\d{2}\.\d{2}\.\d{4})\s+([A-Za-z0-9_]+)\s+(.*)$')

def to_amount(s):
    return float(s.replace('\u00a0','').replace(' ','').replace(',','.'))

recs, cur = [], None
for line in RAW.splitlines():
    line = line.strip()
    if not line: continue
    m = DATE_RE.match(line)
    if m:
        date, code, rest = m.groups()
        am = re.search(r'(-?[\d\s\u00a0]*\d,\d{2})\s*RUR\s*$', rest)
        amt = to_amount(am.group(1)) if am else None
        desc = rest[:am.start()].strip() if am else rest
        cur = {'date': date, 'code': code, 'desc': desc, 'amount': amt}
        recs.append(cur)
    elif cur is not None:
        am = re.search(r'^(-?[\d\s\u00a0]*\d,\d{2})\s*RUR\s*(.*)$', line)
        if am and cur['amount'] is None:
            cur['amount'] = to_amount(am.group(1))
            cur['desc'] += ' ' + am.group(2).strip()
        else:
            cur['desc'] += ' ' + line

skip = re.compile(r'(Уполномоченное лицо|подпись сотрудника|Страница \d+ из|Дата проводки|в валюте счета|Выписка по счету|За период с|Номер счета|Дата открытия|Валюта счета|Тип счета|Дата формирования|Исходящий остаток|Клиент |Платежный лимит|Адрес регистрации|Неподтвержденные|На дату формирования|Текущий баланс|Общая задолженность|Операции по счету|р-н, Щелково|область|Московская)')
recs = [r for r in recs if r['amount'] is not None and not skip.search(r['desc'])]

out=[]
for r in recs:
    if r['amount'] <= 0: continue
    d=r['desc']
    phone=''
    mp=re.search(r'\+7\s*\((\d{3})\)\s*(\d{3})-(\d{2})-\s*(\d{2})?', d)
    if mp:
        digits=''.join(g for g in mp.groups() if g)
        if len(digits)==10: phone='7'+digits
    if not phone:
        mp=re.search(r'от \+?7?(\d{10})(?!\d)', d)
        if mp: phone='7'+mp.group(1)
    mn=re.search(r'[Оо]тправителя ([^.]+?)\.', d)
    sender=mn.group(1).strip() if mn else ''
    mpur=re.search(r'Без НДС\.?\s*(.*)$', d)
    purpose=mpur.group(1).strip() if mpur else ''
    kind=('SBP' if 'Систему быстрых платежей' in d else
          'SBER_phone' if 'по номеру телефона' in d else
          'transfer' if 'Перевод денежных средств' in d else
          'internal' if 'между счетами' in d else 'other')
    out.append({'date':r['date'],'amount':r['amount'],'phone':phone,'sender':sender,
                'kind':kind,'purpose':purpose,'desc':d})

json.dump(out, open('out/alfa_incoming.json','w',encoding='utf-8'), ensure_ascii=False, indent=1)
print('входящих:',len(out),'| с телефоном:',sum(1 for x in out if x['phone']),
      '| с именем:',sum(1 for x in out if x['sender']),
      '| без ключей:',sum(1 for x in out if not x['phone'] and not x['sender']))
for x in out:
    if not x['phone'] and not x['sender']:
        print('  NOKEY:',x['date'],x['amount'],x['kind'],'|',x['desc'][:80])
