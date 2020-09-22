<div align="center">

## Create Dir \(Folder\)


</div>

### Description

this code creates folder with any number of subfolders beneath it. MkDir can't do this!

sorry 4 my english :)
 
### More Info
 
Path as string

paste this code in new module

in immediate window type for example

CreateDir "c:\rrr\ggg\jjj\kkk"


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bibikoff Leonid](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bibikoff-leonid.md)
**Level**          |Intermediate
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bibikoff-leonid-create-dir-folder__1-36942/archive/master.zip)





### Source Code

```
Sub CreateDir(strPath As String)
On Error Resume Next
Dim ArrFolders As Variant
ArrFolders = Split(strPath, "\")
dim i as long
Dim CurPath As String: CurPath = ArrFolders(0)
MkDir CurPath
For i = 1 To UBound(ArrFolders)
  CurPath = CurPath & "\" & ArrFolders(i)
  MkDir CurPath
Next i
On Error GoTo 0
If Len(Dir(strPath, vbDirectory)) = 0 Then
  Err.Raise vbObjectError, , "Can't create dir" & vbCrLf & strPath & vbcrlf & ":(((("
End If
End Sub
```

