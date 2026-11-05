"""Builds the test fixtures with openpyxl (an implementation independent of our reader), then
injects a fake macro part so the byte-preserving save can be proved. Not part of the plugin."""
import os
import zipfile
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURES = os.path.join(os.path.dirname(HERE), "test", "fixtures")
os.makedirs(FIXTURES, exist_ok=True)

thin = Side(style="thin", color="FFB0B0B0")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
head_fill = PatternFill("solid", fgColor="FF1F4E79")
head_font = Font(bold=True, size=11, color="FFFFFFFF")
band = PatternFill("solid", fgColor="FFF2F5F8")

wb = Workbook()
ws = wb.active
ws.title = "Summary"

cols = ["Store", "Sales", "Revenue", "Margin", "Opened"]
widths = [44, 13, 20, 10, 14]
for i, (name, width) in enumerate(zip(cols, widths), start=1):
    c = ws.cell(row=1, column=i, value=name)
    c.fill = head_fill
    c.font = head_font
    c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    c.border = box
    ws.column_dimensions[get_column_letter(i)].width = width

ws.row_dimensions[1].height = 30

rows = [
    ("5063GS_Krasn_1", 282349, 13755522, 0.0812, "2019-04-01"),
    ("4047M_Lazorevy", 271219, 12397811, 0.0640, "2020-11-15"),
    ("4835DS_Lefortovo", 262386, 38821200, 0.0300, "2018-06-30"),
    ("5005DS_Kuzovlevo", 242698, 30156330, 0.0500, "2021-02-12"),
    ("5818DS_Olshansk", 237344, 30887166, 0.0600, "2017-09-08"),
]
for r, (store, sales, revenue, margin, opened) in enumerate(rows, start=2):
    ws.cell(row=r, column=1, value=store).border = box
    c = ws.cell(row=r, column=2, value=sales)
    c.number_format = "#,##0"
    c.border = box
    c.alignment = Alignment(horizontal="right")
    c = ws.cell(row=r, column=3, value=f"=B{r}*48.7")
    c.number_format = '#,##0.00 "RUB"'
    c.border = box
    c = ws.cell(row=r, column=4, value=margin)
    c.number_format = "0.00%"
    c.border = box
    c = ws.cell(row=r, column=5, value=opened)
    c.number_format = "yyyy-mm-dd"
    c.border = box
    if r % 2:
        for col in range(1, 6):
            ws.cell(row=r, column=col).fill = band

total_row = len(rows) + 2
ws.cell(row=total_row, column=1, value="Total").font = Font(bold=True)
c = ws.cell(row=total_row, column=2, value=f"=SUM(B2:B{total_row - 1})")
c.font = Font(bold=True)
c.number_format = "#,##0"
ws.cell(row=total_row, column=3, value=f"=SUM(C2:C{total_row - 1})").number_format = "#,##0.00"

ws.merge_cells(start_row=total_row + 2, start_column=1, end_row=total_row + 2, end_column=5)
m = ws.cell(row=total_row + 2, column=1, value="Prepared from the store ledger")
m.alignment = Alignment(horizontal="center")

ws.freeze_panes = "B2"
ws.conditional_formatting.add(
    f"D2:D{total_row - 1}",
    CellIsRule(operator="lessThan", formula=["0.04"], fill=PatternFill("solid", fgColor="FFFFC7CE")),
)
wb.defined_names.add(DefinedName("SalesTotal", attr_text=f"Summary!$B${total_row}"))

notes = wb.create_sheet("Notes")
notes["A1"] = "Sheet with no styling"
notes["A3"] = "A longer line of text that should stay readable in a narrow column"
notes.column_dimensions["A"].width = 18

plain = wb.create_sheet("Plain")
for r in range(1, 201):
    plain.cell(row=r, column=1, value=r)
    plain.cell(row=r, column=2, value=f"row {r}")

rich = os.path.join(FIXTURES, "rich.xlsx")
wb.save(rich)
print("wrote", rich, os.path.getsize(rich), "bytes")

# A macro-enabled twin: same parts plus a vbaProject part, so a save must leave it byte-identical.
macro = os.path.join(FIXTURES, "macro.xlsm")
src = zipfile.ZipFile(rich, "r")
with zipfile.ZipFile(macro, "w", zipfile.ZIP_DEFLATED) as out:
    for item in src.infolist():
        data = src.read(item.filename)
        if item.filename == "[Content_Types].xml":
            text = data.decode("utf-8")
            text = text.replace(
                "</Types>",
                '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
            )
            data = text.encode("utf-8")
        out.writestr(item, data)
    out.writestr("xl/vbaProject.bin", bytes(range(256)) * 40)
print("wrote", macro, os.path.getsize(macro), "bytes")

# A stand-in for a large, heavily formatted real book: 166 rows by 16 columns, a painted header,
# banded rows with borders on every cell, a frozen pane and mixed number formats. It is generated
# because a real customer workbook must never be committed to a public repository.
styled = os.path.join(FIXTURES, "styled.xlsx")
wb2 = Workbook()
ws2 = wb2.active
ws2.title = "Ledger"

head = ["Store", "Sales", "Revenue", "Plan", "Delta", "Margin", "Staff", "Area"] + [
    f"W{index}" for index in range(1, 9)
]
for column, name in enumerate(head, start=1):
    cell = ws2.cell(row=1, column=column, value=name)
    cell.fill = head_fill
    cell.font = head_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = box
    ws2.column_dimensions[get_column_letter(column)].width = 43 if column == 1 else 14

ws2.row_dimensions[1].height = 30

for row in range(2, 167):
    ws2.cell(row=row, column=1, value=f"{4000 + row}_{('Store ' + str(row)):0<20}"[:40]).border = box
    ws2.cell(row=row, column=2, value=200000 + row * 137).number_format = "#,##0.00"
    ws2.cell(row=row, column=3, value=f"=B{row}*48.7").number_format = "#,##0.00"
    ws2.cell(row=row, column=4, value=250000 + row * 100).number_format = "#,##0"
    ws2.cell(row=row, column=5, value=f"=B{row}-D{row}").number_format = "#,##0;[Red]-#,##0"
    ws2.cell(row=row, column=6, value=(row % 15) / 100).number_format = "0.00%"
    ws2.cell(row=row, column=7, value=5 + row % 7)
    ws2.cell(row=row, column=8, value=80 + row % 40).number_format = "#,##0.0"
    for index in range(9, 17):
        ws2.cell(row=row, column=index, value=(row * index) % 997).number_format = "#,##0"
    for column in range(1, 17):
        ws2.cell(row=row, column=column).border = box
        if row % 2 == 0:
            ws2.cell(row=row, column=column).fill = band

ws2.freeze_panes = "B2"
wb2.save(styled)
print("wrote", styled, os.path.getsize(styled), "bytes")
