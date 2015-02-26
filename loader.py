__author__ = 'ezaurum'

import xlrd
import sys
import codecs

input_file = sys.argv[1]

book = xlrd.open_workbook(input_file)
for n in range(book.nsheets):
    sheet = book.sheet_by_index(n)
    outFile = codecs.open('%s.sql' % sheet.name, 'w', encoding='utf-8')
    columns = []
    for i in range(sheet.nrows):
        if 0 == i:
            cn = []
            for j in sheet.row_values(i):
                cn.append('%s' % j)
            outFile.write('DELETE FROM %s;\nINSERT INTO %s (%s) VALUES ' % (sheet.name, sheet.name, ','.join(cn)))
            continue

        cl = []
        for j in sheet.row_values(i):
            if type(j) is str:
                cl.append("N'%s'" % j.replace("'", "''").replace("\n", "\\n"))
            else:
                cl.append('%s' % j)
        columns.append("(%s)" % ",".join(cl))
    outFile.write(',\n'.join(columns))
