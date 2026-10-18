Attribute VB_Name = "ModFile"
Option Explicit

'==================================================
' modFile Rev1.8 Fix8
' File Information UDF Library
'==================================================


'==================================================
' Common
'==================================================

Private Function GetFSO() As Object

Static fso As Object

If fso Is Nothing Then

Set fso = CreateObject("Scripting.FileSystemObject")

End If

Set GetFSO = fso

End Function



Private Function GetText(ByVal v As Variant) As String

If IsObject(v) Then Exit Function

If IsArray(v) Then Exit Function

If IsError(v) Then Exit Function

If IsEmpty(v) Then Exit Function


GetText = CStr(v)

End Function



Private Function GetInputText(ByVal Target As Variant) As String

If IsObject(Target) Then

If TypeName(Target) = "Range" Then

If Target.Cells.Count = 1 Then

GetInputText = GetText(Target.Value)

End If

End If

Exit Function

End If


GetInputText = GetText(Target)

End Function



Private Function IsHelpRequest(ByVal Target As Variant) As Boolean

IsHelpRequest = (GetInputText(Target) = "?")

End Function



'==================================================
' Input
'==================================================

Private Function IsMultiInput(ByVal Target As Variant) As Boolean

If TypeName(Target) = "Range" Then

IsMultiInput = (Target.Cells.Count > 1)

Else

IsMultiInput = False

End If

End Function



'==================================================
' Path
'==================================================

Private Function NormalizeFolderPath(ByVal FolderPath As String) As String

If FolderPath = "" Then Exit Function


FolderPath = Replace(FolderPath, "\", "/")


Do While Right(FolderPath, 1) = "/"

FolderPath = Left(FolderPath, Len(FolderPath) - 1)

Loop


NormalizeFolderPath = FolderPath & "/"

End Function



Private Function GetFileFolderPath(ByVal f As Object) As String

GetFileFolderPath = NormalizeFolderPath(f.ParentFolder.path)

End Function

'==================================================
' File Object
'==================================================

Private Function GetFile(ByVal FilePath As String) As Object

Dim fso As Object
Dim path As String


Set fso = GetFSO()


path = Replace(FilePath, "/", "\")



If Not fso.FileExists(path) Then Exit Function


Set GetFile = fso.GetFile(path)

End Function



Private Function GetFolder(ByVal FolderPath As String) As Object

Dim fso As Object
Dim path As String


Set fso = GetFSO()


path = Replace(FolderPath, "/", "\")



If Not fso.FolderExists(path) Then Exit Function


Set GetFolder = fso.GetFolder(path)

End Function



'==================================================
' Help Array
'==================================================

Private Function HelpArray(ByVal Text As String) As Variant

Dim Lines() As String
Dim Result() As Variant
Dim i As Long


Lines = Split(Text, vbLf)


ReDim Result(1 To UBound(Lines) + 1, 1 To 1)



For i = 0 To UBound(Lines)

Result(i + 1, 1) = Lines(i)

Next



HelpArray = Result

End Function



'==================================================
' FileInfo Help
'==================================================

Private Function HelpFileInfo() As Variant

Dim Text As String


Text = _
"Function : FileInfo" & vbLf & _
"Description : Returns file information." & vbLf & _
"Syntax : =FileInfo(FilePath)" & vbLf & _
vbLf & _
"Input :" & vbLf & _
"Cell reference supported." & vbLf & _
"UDF result supported." & vbLf & _
"Multiple cells range supported." & vbLf & _
"Blank cells ignored." & vbLf & _
vbLf & _
"Output :" & vbLf & _
"Dynamic Array." & vbLf & _
"Header included." & vbLf & _
"Vertical spill." & vbLf & _
vbLf & _
"Version : Ver2.0"



HelpFileInfo = HelpArray(Text)

End Function



'==================================================
' FileInfo Single
'==================================================

Private Function FileInfoSingle(ByVal FilePath As String) As Variant

Dim Result(1 To 2, 1 To 6) As Variant
Dim f As Object


Result(1, 1) = "Name"
Result(1, 2) = "Path"
Result(1, 3) = "Date"
Result(1, 4) = "Size"
Result(1, 5) = "Ext"
Result(1, 6) = "Exists"



Set f = GetFile(FilePath)



If f Is Nothing Then


Result(2, 6) = False



Else


Result(2, 1) = f.Name
Result(2, 2) = GetFileFolderPath(f)
Result(2, 3) = f.DateLastModified
Result(2, 4) = f.Size
Result(2, 5) = GetFSO().GetExtensionName(f.Name)
Result(2, 6) = True



End If



FileInfoSingle = Result


End Function

'==================================================
' FileInfo Multi
'==================================================

Private Function FileInfoMulti(ByVal Target As Range) As Variant

Dim Result() As Variant
Dim temp As Variant
Dim cell As Range
Dim Row As Long
Dim col As Long
Dim Count As Long
Dim Value As String



For Each cell In Target.Cells


If GetText(cell.Value) <> "" Then

Count = Count + 1

End If


Next



If Count = 0 Then

FileInfoMulti = CVErr(xlErrNA)

Exit Function

End If



ReDim Result(1 To Count + 1, 1 To 6)



Result(1, 1) = "Name"
Result(1, 2) = "Path"
Result(1, 3) = "Date"
Result(1, 4) = "Size"
Result(1, 5) = "Ext"
Result(1, 6) = "Exists"



Row = 2



For Each cell In Target.Cells


Value = GetText(cell.Value)



If Value <> "" Then


temp = FileInfoSingle(Value)



For col = 1 To 6

Result(Row, col) = temp(2, col)

Next



Row = Row + 1



End If



Next



FileInfoMulti = Result


End Function



'==================================================
' FileInfo Public
'==================================================

Public Function FileInfo(ByVal Target As Variant) As Variant


If IsHelpRequest(Target) Then


FileInfo = HelpFileInfo()


Exit Function



End If



If IsMultiInput(Target) Then


FileInfo = FileInfoMulti(Target)



Else


FileInfo = FileInfoSingle(GetInputText(Target))



End If



End Function



'==================================================
' FileList
'==================================================

Private Function FileListCore(ByVal FolderPath As String) As Variant

Dim folder As Object
Dim file As Object
Dim Result() As Variant
Dim Row As Long
Dim path As String



path = NormalizeFolderPath(FolderPath)



Set folder = GetFolder(FolderPath)



If folder Is Nothing Then


FileListCore = CVErr(xlErrNA)


Exit Function



End If



ReDim Result(1 To folder.Files.Count + 1, 1 To 2)



Result(1, 1) = "Name"
Result(1, 2) = "Path"



Row = 2



For Each file In folder.Files


Result(Row, 1) = file.Name

Result(Row, 2) = path


Row = Row + 1



Next



FileListCore = Result


End Function

'==================================================
' FileList Help
'==================================================

Private Function HelpFileList() As Variant

Dim Text As String


Text = _
"Function : FileList" & vbLf & _
"Description : Returns file list." & vbLf & _
"Syntax : =FileList(FolderPath)" & vbLf & _
vbLf & _
"Input :" & vbLf & _
"Cell reference supported." & vbLf & _
"UDF result supported." & vbLf & _
"Multiple cells range not supported." & vbLf & _
vbLf & _
"Output :" & vbLf & _
"Dynamic Array." & vbLf & _
"Header included." & vbLf & _
"Vertical spill." & vbLf & _
vbLf & _
"Version : Ver2.0"



HelpFileList = HelpArray(Text)

End Function



'==================================================
' FileList Public
'==================================================

Public Function FileList(ByVal Target As Variant) As Variant


If IsHelpRequest(Target) Then


FileList = HelpFileList()


Exit Function



End If



FileList = FileListCore(GetInputText(Target))


End Function



'==================================================
' Single Function Help
'==================================================

Private Function HelpSingleFile(ByVal FunctionName As String, _
ByVal Description As String, _
ByVal Syntax As String, _
ByVal Output As String) As Variant

Dim Text As String


Text = _
"Function : " & FunctionName & vbLf & _
"Description : " & Description & vbLf & _
"Syntax : " & Syntax & vbLf & _
vbLf & _
"Input :" & vbLf & _
"Cell reference supported." & vbLf & _
"UDF result supported." & vbLf & _
vbLf & _
"Output :" & vbLf & _
Output & vbLf & _
vbLf & _
"Version : Ver2.0"



HelpSingleFile = HelpArray(Text)

End Function



'==================================================
' FileExists
'==================================================

Public Function FileExists(ByVal Target As Variant) As Variant


If IsHelpRequest(Target) Then


FileExists = HelpSingleFile( _
"FileExists", _
"Checks whether file exists.", _
"=FileExists(FilePath)", _
"TRUE / FALSE")


Exit Function


End If



FileExists = GetFSO().FileExists(Replace(GetInputText(Target), "/", "\"))


End Function



'==================================================
' FileName
'==================================================

Public Function FileName(ByVal Target As Variant) As Variant

Dim f As Object


If IsHelpRequest(Target) Then


FileName = HelpSingleFile( _
"FileName", _
"Returns file name.", _
"=FileName(FilePath)", _
"File name")


Exit Function


End If



Set f = GetFile(GetInputText(Target))


If f Is Nothing Then Exit Function



FileName = f.Name


End Function

'==================================================
' FilePath
'==================================================

Public Function FilePath(ByVal Target As Variant) As Variant

Dim f As Object


If IsHelpRequest(Target) Then


FilePath = HelpSingleFile( _
"FilePath", _
"Returns folder path of file.", _
"=FilePath(FilePath)", _
"Folder path ending with /")


Exit Function


End If



Set f = GetFile(GetInputText(Target))


If f Is Nothing Then Exit Function



FilePath = GetFileFolderPath(f)


End Function



'==================================================
' FileExt
'==================================================

Public Function FileExt(ByVal Target As Variant) As Variant

Dim f As Object


If IsHelpRequest(Target) Then


FileExt = HelpSingleFile( _
"FileExt", _
"Returns file extension.", _
"=FileExt(FilePath)", _
"File extension")


Exit Function


End If



Set f = GetFile(GetInputText(Target))


If f Is Nothing Then Exit Function



FileExt = GetFSO().GetExtensionName(f.Name)


End Function



'==================================================
' FileSize
'==================================================

Public Function FileSize(ByVal Target As Variant) As Variant

Dim f As Object


If IsHelpRequest(Target) Then


FileSize = HelpSingleFile( _
"FileSize", _
"Returns file size.", _
"=FileSize(FilePath)", _
"File size in bytes")


Exit Function


End If



Set f = GetFile(GetInputText(Target))


If f Is Nothing Then Exit Function



FileSize = f.Size


End Function



'==================================================
' FileDate
'==================================================

Public Function FileDate(ByVal Target As Variant) As Variant

Dim f As Object


If IsHelpRequest(Target) Then


FileDate = HelpSingleFile( _
"FileDate", _
"Returns file modified date.", _
"=FileDate(FilePath)", _
"Last modified date")


Exit Function


End If



Set f = GetFile(GetInputText(Target))


If f Is Nothing Then Exit Function



FileDate = f.DateLastModified


End Function



