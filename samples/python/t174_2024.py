# created by FrancescoSatriani98 in https://github.com/testo/toolbox_classic/issues/4
import win32com.client
inst = win32com.client.Dispatch("tcddka.tcddk")
com_port = 6
initOk = inst.Init(com_port-1, "testo174-2024", 4000)
if initOk:
  numCol = inst.NumCols
  inst.Get
  for i in range(numCol):
      if inst.IsNonNumeric(i):
          print(inst.StringVal(i))
      else:
          print(f"{inst.Val(i)} {inst.Unit(i)}")