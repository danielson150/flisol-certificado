# -*- coding: utf-8 -*-
from lxml import etree
from subprocess import Popen
import xlrd
import os

SVGNS = u"http://www.w3.org/2000/svg"

with open('./certificado/certificado.svg', 'r') as mysvg:
    svg = mysvg.read()

svg = str(svg)

xml_data = etree.fromstring(svg)
# We search for element 'text' with id='tile_text' in SVG namespace
find_text = etree.ETXPath("//{%s}tspan[@id='tspan3951']" % (SVGNS))
# find_text(xml_data) returns a list
# [<Element {http://www.w3.org/2000/svg}text at 0x106185ab8>]
# take the 1st element from the list, replace the text

workbook = xlrd.open_workbook('./speakers/speakers.xls')
worksheet = workbook.sheet_by_name('Speakers')
for x in range(1,15,1):
    id = str(int(worksheet.cell(x,0).value))
    name = ''
    for y in range(1,4,1):
        name += worksheet.cell(x, y).value
	name += ' '
    name.rstrip()
    find_text(xml_data)[0].text = unicode(name)
    new_svg = etree.tostring(xml_data)
    cm = 'touch ./speakers/' + id + '.svg'
    os.system(cm)
    f = open( './speakers/' + id + '.svg', 'w' )
    f.write( new_svg )
    f.close()
    svg_file = './speakers/' + id + '.svg'
    pdf_file = './speakers/' + id + '.pdf'
    x = Popen(['/usr/bin/inkscape', svg_file, \
        '--export-pdf=%s' % pdf_file])
