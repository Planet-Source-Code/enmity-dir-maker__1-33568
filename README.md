<div align="center">

## Dir Maker


</div>

### Description

Make any level directories, such as:

MakeDir "c:\abc\1234\aaaa\1111\6666\ggggggggg\dddddddddd\ssssssss\7676\dsdsds"
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[enmity](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/enmity.md)
**Level**          |Advanced
**User Rating**    |3.8 (19 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/enmity-dir-maker__1-33568/archive/master.zip)





### Source Code

```
Public Function MakeDir(Path As String) As Boolean
On Error Resume Next
    Dim o_strRet As String
    Dim o_intItems As Integer
    Dim o_vntItem As Variant
    Dim o_strItems() As String
    o_strItems() = Split(Path, "\")
    o_intItems = 0
    For Each o_vntItem In o_strItems()
      o_intItems = o_intItems + 1
      If o_intItems = 1 Then
        o_strRet = o_vntItem
      Else
        o_strRet = o_strRet & "\" & o_vntItem
        MkDir o_strRet
      End If
    Next
    MakeDir = (Err.Number = 0)
End Function
```

