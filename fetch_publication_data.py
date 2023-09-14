import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

beamlines = {
    '2-ID':    ['Soft Inelastic X-ray Scattering','SIX'],
    '3-ID':    ['Hard X-ray Nanoprobe','HXN'],
    '4-BM':    ['X-ray Fluorescence Microscopy','XFM'],
    '4-ID':    ['Integrated In situ and Resonant Hard X-ray Studies','ISR'],
    '5-ID':    ['Submicron Resolution X-ray Spectroscopy','SRX'],
    '6-BM':    ['Beamline for Materials Measurement','BMM'],
    '7-BM':    ['Quick x-ray Absorption and Scattering','QAS'],
    '7-ID-1':  ['Spectroscopy Soft and Tender','SST1'],
    '7-ID-2':  ['Spectroscopy Soft and Tender 2','SST2'],
    '8-BM':    ['Tender Energy X-ray Absorption Spectroscopy','TES'],
    '8-ID':    ['Inner-Shell Spectroscopy','ISS'],
    '10-ID':   ['Inelastic X-ray Scattering','IXS'],
    '11-BM':   ['Complex Materials Scattering','CMS'],
    '11-ID':   ['Coherent Hard X-ray Scattering','CHX'],
    '12-ID':   ['Soft Matter Interfaces','SMI'],
    '16-ID':   ['Life Science X-ray Scattering','LiX'],
    '17-BM':   ['X-ray Footprinting of Biological Materials','XFP'],
    '17-ID-1': ['Highly Automated Macromolecular Crystallography','AMX'],
    '17-ID-2': ['Frontier Microfocusing Macromolecular Crystallography','FMX'],
    '18-ID':   ['Full Field X-ray Imaging','FXI'],
    '19-ID':   ['Biological Microdiffraction Facility','NYX'],
    '21-ID':   ['Electron Spectro-Microscopy','ESM'],
    '22-IR-1': ['Frontier Synchrotron Infrared Spectroscopy','FIS'],
    '22-IR-2': ['Magnetospectroscopy, Ellipsometry and Time-Resolved Optical Spectroscopie','MET'],
    '23-ID-1': ['Coherent Soft X-ray Scattering beamline','CSX'],
    '23-ID-2': ['In situ and Operando Soft X-ray Spectroscopy','IOS'],
    '27-ID':   ['High Energy Engineering X-ray Scattering','HEX'],
    '28-ID-1': ['Pair Distribution Function','PDF'],
    '28-ID-2': ['X-ray Powder Diffraction','XPD'],
}

wb = Workbook()
ws = wb.active
ws.append(['port', 'TLA', 'total', 'high profile', 'years operating',
           '% high profile', 'pubs per year'])
c = ws['A2']
ws.freeze_panes = c
for cell in range(20):
    ws[f'{chr(65+cell)}1'].font = Font(bold=True)

for i, port in enumerate(beamlines.keys()):
    tla = beamlines[port][1]
    url = f'https://www.bnl.gov/nsls2/beamlines/publications.php?q={port}'
    html = requests.get(url).content
    df_list = pd.read_html(html)
    df = df_list[-1]
    answer = str(df).split('\n')[-1].split()
    years = len(str(df).split('\n')) - 2
    ws.append([port, tla, answer[3], answer[2], years,
               f'=D{i+3}/C{i+3}',
               f'=C{i+3}/E{i+3}',
               ])
    ws[f'F{i+3}'].number_format = '0.00'
    ws[f'G{i+3}'].number_format = '0.00'


ws.column_dimensions['D'].width = 1.2 *  ws.column_dimensions['D'].width
ws.column_dimensions['E'].width = 1.3 *  ws.column_dimensions['E'].width
ws.column_dimensions['F'].width = 1.2 *  ws.column_dimensions['F'].width
ws.column_dimensions['G'].width = 1.2 *  ws.column_dimensions['G'].width
wb.save("beamline_publication_data.xlsx")
