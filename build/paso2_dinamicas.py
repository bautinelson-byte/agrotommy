"""Paso 2: crea las tablas dinámicas nativas en 'Tabla Resumen' con LibreOffice (UNO) y guarda .xlsx."""
import os
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.table import CellAddress, CellRangeAddress

SRC, DST = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
NR = int(sys.argv[3])  # última fila de datos de Estructura EV

ROW, COL, PAGE, DATA = (uno.Enum('com.sun.star.sheet.DataPilotFieldOrientation', n)
                        for n in ('ROW', 'COLUMN', 'PAGE', 'DATA'))
SUM = uno.Enum('com.sun.star.sheet.GeneralFunction', 'SUM')
C = {'ctr': 0, 'campo': 1, 'cuenta': 2, 'art': 3, 'pesos': 12, 'anio': 13, 'mes': 14, 'usd': 16}


def pv(n, v):
    p = PropertyValue()
    p.Name, p.Value = n, v
    return p


import glob
sys.path.insert(0, glob.glob('/root/.claude/skills/synced/*/xlsx/scripts')[0])
from office.soffice import get_soffice_env  # noqa: E402

proc = subprocess.Popen(['soffice', '--headless', '--invisible', '--norestore',
                         '--accept=socket,host=localhost,port=2002;urp;'], env=get_soffice_env())
ctx = None
for _ in range(60):
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        ctx = resolver.resolve('uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext')
        break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(SRC), '_blank', 0, (pv('Hidden', True), pv('FilterName', 'Calc MS Excel 2007 XML')))
doc.calculateAll()
sheets = doc.Sheets
src_idx = [sheets.getByIndex(i).Name for i in range(sheets.Count)].index('Estructura EV')
tr = sheets.getByName('Tabla Resumen')
tr_idx = [sheets.getByIndex(i).Name for i in range(sheets.Count)].index('Tabla Resumen')
dps = tr.getDataPilotTables()

# (título, filas, columnas, filtro de página, ocultos {campo: [items]} )
SIN_GAN = {'ctr': ['Ganaderia']}
SOLO_GAN = {'ctr': ['Administración', 'Estructura']}
TABLAS = [
    ('Gastos por cuenta y campo (Estructura + Administración)', ['cuenta'], ['campo'], SIN_GAN),
    ('Gastos mensuales por cuenta (Estructura + Administración)', ['cuenta'], ['campo', 'anio', 'mes'], SIN_GAN),
    ('Impuestos mensuales por artículo', ['cuenta', 'art'], ['campo', 'anio', 'mes'], dict(SIN_GAN, cuenta='SOLO:Impuestos')),
    ('Gastos por cuenta y artículo (Estructura + Administración)', ['cuenta', 'art'], ['campo'], SIN_GAN),
    ('Ganadería por cuenta', ['cuenta'], ['campo'], SOLO_GAN),
    ('Ganadería mensual', ['cuenta'], ['anio', 'mes', 'campo'], SOLO_GAN),
    ('Sueldos por campo', ['cuenta', 'art'], ['campo'], dict(SIN_GAN, cuenta='SOLO:Sueldos')),
    ('Sueldos mensual', ['cuenta', 'art'], ['anio', 'mes', 'campo'], dict(SIN_GAN, cuenta='SOLO:Sueldos')),
]

row = 2
tr.getCellByPosition(0, 0).setString('Tablas dinámicas sobre Estructura EV. Actualizar: Datos > Actualizar todo. '
                                     'Cada tabla está en pesos y debajo en dólares (BCRA fin de mes).')
n = 0
for titulo, filas, cols, ocultos in TABLAS:
    for moneda, campo_dato in (('$', 'pesos'), ('USD', 'usd')):
        n += 1
        desc = dps.createDataPilotDescriptor()
        desc.setSourceRange(CellRangeAddress(src_idx, 0, 0, 17, NR - 1))
        flds = desc.getDataPilotFields()
        f = flds.getByIndex(C['ctr'])
        f.Orientation = PAGE
        for pos, k in enumerate(filas):
            f = flds.getByIndex(C[k])
            f.Orientation = ROW
        for k in cols:
            flds.getByIndex(C[k]).Orientation = COL
        d = flds.getByIndex(C[campo_dato])
        d.Orientation = DATA
        d.Function = SUM
        for k, val in ocultos.items():
            items = flds.getByIndex(C[k]).getItems()
            names = items.getElementNames()
            for nm in names:
                hide = (nm != val[5:]) if isinstance(val, str) else (nm in val)
                if hide:
                    items.getByName(nm).IsHidden = True
        title = f'{titulo} - {moneda}'
        tr.getCellByPosition(0, row).setString(title)
        tr.getCellByPosition(0, row).CharWeight = 150
        name = f'Tabla{n:02d}'
        dps.insertNewByName(name, CellAddress(tr_idx, 0, row + 1), desc)
        out = dps.getByName(name).getOutputRange()
        print(name, title, f'filas {out.StartRow + 1}-{out.EndRow + 1}', f'cols {out.StartColumn}-{out.EndColumn}')
        row = out.EndRow + 4

doc.calculateAll()
doc.storeToURL(uno.systemPathToFileUrl(DST), (pv('FilterName', 'Calc MS Excel 2007 XML'),))
doc.close(True)
try:
    desktop.terminate()
except Exception:
    pass
proc.wait(timeout=60)
