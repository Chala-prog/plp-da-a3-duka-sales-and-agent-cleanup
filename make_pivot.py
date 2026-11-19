"""Uses headless LibreOffice (UNO) to add a real PivotTable (DataPilot) to a workbook that has a 'Data' sheet."""
import subprocess, time, sys, os, uno
from com.sun.star.table import CellRangeAddress, CellAddress
from com.sun.star.beans import PropertyValue
from com.sun.star.sheet.DataPilotFieldOrientation import ROW, COLUMN, DATA
from com.sun.star.sheet.GeneralFunction import SUM

src, dst, last_row = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2]), int(sys.argv[3])
p = subprocess.Popen(["soffice","--headless","--norestore","--accept=socket,host=localhost,port=2002;urp;"],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
ctx = None
for _ in range(40):
    try:
        local = uno.getComponentContext()
        res = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        ctx = res.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(1)
smgr = ctx.ServiceManager; desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
pv = lambda n,v: PropertyValue(Name=n, Value=v)
doc = desktop.loadComponentFromURL("file://"+src, "_blank", 0, (pv("Hidden",True),))
sheets = doc.Sheets
sheets.insertNewByName("Pivot_Branch_Channel", 1)
data = sheets.getByName("Data"); out = sheets.getByName("Pivot_Branch_Channel")
hdr = [data.getCellByPosition(c,0).String for c in range(8)]
dp = out.DataPilotTables
desc = dp.createDataPilotDescriptor()
desc.SourceRange = CellRangeAddress(Sheet=data.RangeAddress.Sheet, StartColumn=0, StartRow=0, EndColumn=7, EndRow=last_row-1)
f = desc.DataPilotFields
f.getByName("branch").Orientation = ROW
f.getByName("channel").Orientation = COLUMN
d = f.getByName("revenue"); d.Orientation = DATA; d.Function = SUM
dp.insertNewByName("PivotTable1", CellAddress(Sheet=out.RangeAddress.Sheet, Column=0, Row=2), desc)
doc.storeToURL("file://"+dst, (pv("FilterName","Calc MS Excel 2007 XML"),))
doc.close(True)
try: desktop.terminate()
except Exception: pass
p.terminate()
