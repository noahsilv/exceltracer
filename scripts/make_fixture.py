"""Optional fixture regeneration. Requires XlsxWriter and the installed npm development tools."""
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
import xlsxwriter

root = Path(__file__).resolve().parents[1]
memory = BytesIO()
book = xlsxwriter.Workbook(memory)
header = book.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#155541'})
ratio = book.add_format({'num_format': '0.00"x"'})
sheets = {name: book.add_worksheet(name) for name in ('Summary', 'Cash Flow', 'Debt')}
for sheet in sheets.values():
    sheet.set_column('A:A', 32)
    sheet.set_column('B:B', 16)
    sheet.write_row('A1', ['Metric', 'Year 1'], header)
summary, cash, debt = sheets.values()
summary.write('A2', 'Debt service coverage ratio')
summary.write_formula('B2', "='Cash Flow'!B4/Debt!B4", ratio, 1.5)
summary.set_selection('B2')
cash.write_column('A2', ['Operating cash receipts', 'Operating costs', 'Cash available for debt service'])
cash.write_column('B2', [200, 50])
cash.write_formula('B4', '=B2-B3', None, 150)
debt.write_column('A2', ['Interest expense', 'Principal repayment', 'Debt service', 'Range trace example'])
debt.write_column('B2', [25, 75])
debt.write_formula('B4', '=B2+B3', None, 100)
debt.write_formula('B5', '=SUM(B2:B3)', None, 100)
forecast = book.add_worksheet('Map example')
forecast.set_column('A:A', 26)
forecast.set_column('B:G', 13)
forecast.write_row('A1', ['Forecast', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'], header)
forecast.write_column('A2', ['Volume', 'Price', 'Revenue'])
forecast.write_row('B2', [10, 12, 14, 16, 18, 20])
forecast.write_row('B3', [5] * 6)
for i, col in enumerate('BCDEFG'):
    forecast.write_formula(col+'4', '='+col+'2*'+col+'3', None, (10+2*i)*5)
forecast.write_number('E4', 80, book.add_format({'bg_color':'#F2D5E0','bold':True}))
forecast.write('A6','Positive revenue test')
forecast.write_formula('B6','=IF(B4>0,B4/B2,0)',None,5)
# Two direct uses of the same DSCR on different sheets make reverse tracing observable.
cash.write('A7', 'Coverage headroom')
cash.write_formula('B7', '=Summary!B2-1', ratio, 0.5)
debt.write('A7', 'Coverage test')
debt.write_formula('B7', '=IF(Summary!B2>=1,"Pass","Fail")', None, 'Pass')
book.close()
# Mac Excel discovers the add-in in its wef catalog. This workbook has no Windows registry binding.
(root / 'examples/Debt-Model-Mac.xlsx').write_bytes(memory.getvalue())

# Attach Microsoft's developer-catalog task pane to the workbook opened by npm start.
template = root / 'node_modules/office-addin-dev-settings/templates/ExcelWorkbookWithTaskPane.xlsx'
manifest = ET.parse(root / 'manifest.xml').getroot()
def manifest_text(tag):
    return manifest.find('{*}' + tag).text
with ZipFile(memory) as base, ZipFile(template) as addin, ZipFile(root / 'examples/Debt-Model.xlsx', 'w', ZIP_DEFLATED) as out:
    for name in base.namelist():
        data = base.read(name)
        if name in ('_rels/.rels', '[Content_Types].xml'):
            xml = ET.fromstring(data)
            for entry in ET.fromstring(addin.read(name)):
                if 'webextension' in str(entry.attrib):
                    xml.append(entry)
            data = ET.tostring(xml, encoding='utf-8', xml_declaration=True)
        out.writestr(name, data)
    for name in addin.namelist():
        if name.startswith('xl/webextensions/'):
            data = addin.read(name)
            if name.endswith('/webextension.xml'):
                data = data.replace(b'00000000-0000-0000-0000-000000000000', manifest_text('Id').encode())
                data = data.replace(b'1.0.0.0', manifest_text('Version').encode())
            out.writestr(name, data)
