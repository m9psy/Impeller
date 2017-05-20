# coding: utf-8

from impeller.workbook import WorkBook

# TEST()

import time
import datetime

ROWS = 10**5
COLS = 10

c_wb = WorkBook("Юникод имя_c.xlsx")
c_wb.set_properties({"author": "Емельянов Дмитрий", "company": "Нету компании"})
c_wb.set_custom_property("Кастом юникод дата", datetime.datetime.now())
c_ws = c_wb.add_worksheet("Юникод шит!")
start = time.clock()
bold_format = c_wb.add_format()
bold_format.set_bold()
for i in range(ROWS):
    for j in range(COLS):
        # row, col, data
        c_ws.write_number(i, j, i + j, None if i else bold_format)
c_ws.set_column(0, 2, 30)
c_ws.set_column("C:E", 50)
start_close = time.clock()
c_wb.close()
end_close = time.clock()
end = time.clock()
print("Close time:", format(end_close - start_close, ".2f"), "Total:", format(end - start, ".2f"))