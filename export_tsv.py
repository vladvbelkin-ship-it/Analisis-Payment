# -*- coding: utf-8 -*-
"""Выгрузка сопоставленных платежей по центрам в TSV.

Формат строки:  Дата \t Имя ребенка \t Имя родителя \t Сумма \t Примечание \t Банк
Примечание содержит критерий совпадения и вероятность (%).
Порог включения в файл центра: prob >= 60 (K1/K2/K3/K4).
"""
import json
from collections import defaultdict

R = json.load(open('out/payments_matched3.json', encoding='utf-8'))
CENTERS = ['Вольная', 'Удальцова']


def fmt_amount(a):
    a = round(a, 2)
    return str(int(a)) if a == int(a) else f'{a:.2f}'


rows_by_center = defaultdict(list)
for r in R:
    if not r['centers'] or r['prob'] < 60:
        continue
    child = '; '.join(r['children']) if r['children'] else ''
    parent = '; '.join(r['parents']) if r['parents'] else ''
    note = r['crit'] + f' — {r["prob"]}%'
    if len(r['centers']) > 1:
        note += ' | клиент есть и в другом центре — требует проверки'
    if r['already']:
        note += ' | ' + '; '.join(r['already'])
    for c in r['centers']:
        rows_by_center[c].append((r['date'], child, parent, fmt_amount(r['amount']),
                                  note, r['bank']))

for c in CENTERS:
    rows = sorted(rows_by_center[c], key=lambda x: (x[5], x[0]))
    fn = f'out/оплаты_{c}.tsv'
    with open(fn, 'w', encoding='utf-8') as f:
        f.write('Дата\tИмя ребенка\tИмя родителя\tСумма\tПримечание\tБанк\n')
        for row in rows:
            f.write('\t'.join(row) + '\n')
    print(c, '->', fn, '| строк:', len(rows),
          '| сумма:', sum(float(x[3]) for x in rows))

# сводка по банкам/вероятностям
st = defaultdict(lambda: defaultdict(int))
for c in CENTERS:
    for row in rows_by_center[c]:
        st[c][(row[5], row[4].split('— ')[-1])] += 1
print()
for c in CENTERS:
    print('===', c)
    for k, v in sorted(st[c].items()):
        print(f'   {k[0]} | {k[1]}: {v}')
