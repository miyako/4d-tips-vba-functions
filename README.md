# 4d-tips-vba-functions
Helper functions for Unicode support in VBA

```vba
Private Function AscU(char As String) As Long
  AscU = VBA.CLng("&H0000" + (VBA.Hex(VBA.AscW(char))))
End Function
```
