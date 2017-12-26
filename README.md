# 4d-tips-vba-functions
Helper functions for Unicode support in VBA

```vba
'AscW returns signed integer, which can be negative for 0x8000 and above
Public Function AscU(char As String) As Long
  AscU = VBA.CLng("&H0000" + (VBA.Hex(VBA.AscW(char))))
End Function
```
