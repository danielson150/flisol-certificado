#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import xlrd
import cairosvg
from PyPDF2 import PdfFileMerger

source = './certificado/certificado.svg'

workbook = xlrd.open_workbook('./speakers/speakers.xls')
worksheet = workbook.sheet_by_name('Speakers')
merger = PdfFileMerger()
for x in range(1,15,1):
    id = str(int(worksheet.cell(x,0).value))
    name = ''
    for y in range(1,4,1):
        name += worksheet.cell(x, y).value
        name += ' '
    name.rstrip()
    svg_file = './speakers/' + id + '.svg'
    open(svg_file, 'w').write(open(source).read().replace('>&lt;Nombre Completo&gt;</tspan>','>' + name + '</tspan>'))
    pdf_file = './speakers/' + id + '.pdf'
    cairosvg.svg2pdf(url=svg_file, write_to=pdf_file)
    merger.append(open(pdf_file, 'rb'))

with open('./speakers/cert.pdf', 'wb') as fout:
    merger.write(fout)
