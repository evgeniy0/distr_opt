# distr_opt
#Put next 2 lines manually
cd C:\Python27
python.exe
#Copy and past all below
import win32com.client as c
rastr = c.Dispatch("Astra.Rastr")
rastr
dir(rastr)
com_object = rastr
rastr.Load(1, "c:\\IEEE14bus_reg.rg2", "c:\\IEEE14bus_reg.rg2")
rastr.Tables
tables = rastr.Tables
dir(tables)
dir(rastr.Load)
tables.Find("node")
# expected '0' output
Node = rastr.Tables('node')
tvetv = rastr.Tables('vetv')
print('расчет (проверка) режима для оптимизации (cor-файла)')
kod = rastr.rgm('')
if kod !=0: print('ERROR!')
rastr.opt('')
tvetv.SetSel('sta')
tvetv.WriteCSV(1, file, 'ip,iq,nameu', ';')
