#!/usr/bin/python3
# -*- coding: utf-8 -*-

import xlrd
import xlwt

inp_file = 'bamby2.xls'
out_file = 'bg_' + inp_file

rb = xlrd.open_workbook(inp_file, formatting_info=True)
sheet = rb.sheet_by_index(0)

font_h0 = xlwt.Font()
font_h0.name = 'Arial'
font_h0.colour_index = 0
font_h0.bold = True
font_h0.height = 12 * 20

font0 = xlwt.Font()
font0.name = 'Arial'
font0.colour_index = 0
font0.bold = True
font0.height = 10 * 20

font1 = xlwt.Font()
font1.name = 'Times New Roman'
font1.colour_index = 0
font1.bold = False
font1.height = 10 * 20

style_h = xlwt.XFStyle()
style_h.font = font0

style_h0 = xlwt.XFStyle()
style_h0.font = font_h0

style = list()
for i in range(7):
    s = xlwt.XFStyle()
    s.font = font1
    style.append(s)

style[1].num_format_str = '#,##0'
style[2].num_format_str = '#,##0.0'

wb = xlwt.Workbook()
ws = wb.add_sheet('TDSheet')

Header = ['Номенклатура', 'Кол-во', 'Цена', 'Заказ', 'Штрих', 'ФОТО', 'Сумма заказа']
for idx, str in enumerate(Header):
    ws.write(2, idx, str, style_h0)

cur_header = ''
header_written = False
row_idx = 3

for row in sheet.get_rows():
    c = row[0]
    v = c.value

    if c.xf_index == 74:
        cur_header = v
        header_written = False

    if 'Наст.' in v:
        if not header_written:
            ws.write(row_idx, 0, cur_header, style_h)
            row_idx += 1
            header_written = True
        
        for idx, col in enumerate(row[:6]):
            ws.write(row_idx, idx, col.value, style[idx])
            
        formula = xlwt.Formula('C%d*D%d' % (row_idx+1, row_idx+1))
        ws.write(row_idx, 6, formula, style[6])
        
        row_idx += 1

ws.write(0, 5, 'Сумма заказа', style[0])
ws.write(0, 6, xlwt.Formula('SUM(G%d:G%d)' % (4, row_idx)))
ws.row(2).height = 350
ws.col(0).set_width(20000)
ws.col(4).set_width(4000)
wb.save(out_file)
