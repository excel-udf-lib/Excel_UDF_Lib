Attribute VB_Name = "modCell"
Option Explicit

'==================================================
' Excel UDF Lib Ver2.0
' Module : modCell
' Part1 : Common / CellInfo
'==================================================

'==================================================
' Common
'==================================================

Private Const HELP_MARK As String = "?"

Private Function GetInputText(ByVal Target As Variant) As String

If TypeName(Target) = "Range" Then

If Target.Cells.Count = 1 Then

GetInputText = CStr(Target.Value)

End If

Else

GetInputText = CStr(Target)

End If

End Function

Private Function IsHelp(ByVal Target As Variant) As Boolean

IsHelp = (GetInputText(Target) = HELP_MARK)

End Function

Private Function GetTargetRange(ByVal Target As Variant) As Range

On Error Resume Next

If TypeName(Target) = "Range" Then

Set GetTargetRange = Target

Else

Set GetTargetRange = Application.Range(Target)

End If

On Error GoTo 0

End Function

Private Function ReturnNA() As Variant

ReturnNA = CVErr(xlErrNA)

End Function

Private Function ReturnValueError() As Variant

ReturnValueError = CVErr(xlErrValue)

End Function

'==================================================
' Cell Type
'==================================================

Private Function GetCellType(ByVal Target As Range) As String


If IsError(Target.Value) Then

GetCellType = "Error"


ElseIf Target.HasFormula Then

GetCellType = "Formula"


ElseIf IsEmpty(Target.Value) Then

GetCellType = "Blank"


Else

Select Case VarType(Target.Value)

Case vbString

GetCellType = "String"


Case vbByte, vbInteger, vbLong, vbSingle, vbDouble

GetCellType = "Number"


Case vbDate

GetCellType = "Date"


Case Else

GetCellType = "Other"


End Select


End If


End Function

'==================================================
' Blank
'==================================================

Private Function IsCellBlank(ByVal Target As Range) As Boolean

IsCellBlank = (Target.Value = "")

End Function

'==================================================
' CellInfo
'==================================================

Public Function CellInfo(Optional ByVal Target As Variant) As Variant

Dim rng As Range

If IsHelp(Target) Then

CellInfo = HelpCellInfo()

Exit Function

End If

If IsMissing(Target) Then

Set rng = Application.Caller

Else

Set rng = GetTargetRange(Target)

End If

If rng Is Nothing Then

CellInfo = ReturnValueError()

Exit Function

End If

CellInfo = CellInfoCore(rng)

End Function

'==================================================
' CellInfo Core
'==================================================

Private Function CellInfoCore(ByVal Target As Range) As Variant

Dim Result() As Variant
Dim cell As Range
Dim Row As Long
Dim col As Long

ReDim Result(1 To Target.Cells.Count + 1, 1 To 8)

Result(1, 1) = "Address"
Result(1, 2) = "Value"
Result(1, 3) = "Type"
Result(1, 4) = "NumberFormat"
Result(1, 5) = "HasFormula"
Result(1, 6) = "Formula"
Result(1, 7) = "Blank"
Result(1, 8) = "Error"

Row = 2

For col = 1 To Target.Columns.Count

For Each cell In Target.Columns(col).Cells

Result(Row, 1) = cell.Address(False, False)

If IsEmpty(cell.Value) Then
    Result(Row, 2) = ""
Else
    Result(Row, 2) = cell.Value
End If

Result(Row, 3) = GetCellType(cell)

Result(Row, 4) = cell.NumberFormat

Result(Row, 5) = cell.HasFormula

If cell.HasFormula Then

Result(Row, 6) = cell.Formula

Else

Result(Row, 6) = ""

End If

Result(Row, 7) = IsCellBlank(cell)

Result(Row, 8) = IsError(cell.Value)

Row = Row + 1

Next cell

Next col

CellInfoCore = Result

End Function

'==================================================
' Excel UDF Lib Ver2.0
' Module : modCell
' Part2 : Single Function
'==================================================

'==================================================
' CellType
'==================================================

Public Function CellType(Optional ByVal Target As Variant) As Variant

If IsHelp(Target) Then

CellType = HelpCellType()

Exit Function

End If

CellType = GetCellSingleValue(Target, "Type")

End Function

'==================================================
' CellHasFormula
'==================================================

Public Function CellHasFormula(Optional ByVal Target As Variant) As Variant

If IsHelp(Target) Then

CellHasFormula = HelpCellHasFormula()

Exit Function

End If

CellHasFormula = GetCellSingleValue(Target, "HasFormula")

End Function

'==================================================
' CellFormula
'==================================================

Public Function CellFormula(Optional ByVal Target As Variant) As Variant

If IsHelp(Target) Then

CellFormula = HelpCellFormula()

Exit Function

End If

CellFormula = GetCellSingleValue(Target, "Formula")

End Function

'==================================================
' CellFormat
'==================================================

Public Function CellFormat(ByVal Target As Variant) As Variant

If IsHelp(Target) Then

CellFormat = HelpCellFormat()

Exit Function

End If

CellFormat = GetCellSingleValue(Target, "Format")

End Function

'==================================================
' CellBlank
'==================================================

Public Function CellBlank(ByVal Target As Variant) As Variant

If IsHelp(Target) Then

CellBlank = HelpCellBlank()

Exit Function

End If

CellBlank = GetCellSingleValue(Target, "Blank")

End Function

'==================================================
' CellError
'==================================================

Public Function CellError(ByVal Target As Variant) As Variant

If IsHelp(Target) Then

CellError = HelpCellError()

Exit Function

End If

CellError = GetCellSingleValue(Target, "Error")

End Function

'==================================================
' GetCellSingleValue
'==================================================

Private Function GetCellSingleValue(ByVal Target As Variant, ByVal InfoType As String) As Variant

Dim cell As Range

Set cell = GetTargetRange(Target)

If cell Is Nothing Then

GetCellSingleValue = ReturnValueError()

Exit Function

End If

If cell.Cells.Count <> 1 Then

GetCellSingleValue = ReturnValueError()

Exit Function

End If

Select Case InfoType

Case "Type"

GetCellSingleValue = GetCellType(cell)

Case "HasFormula"

GetCellSingleValue = cell.HasFormula

Case "Formula"

If cell.HasFormula Then

GetCellSingleValue = cell.Formula

Else

GetCellSingleValue = ""

End If

Case "Format"

GetCellSingleValue = cell.NumberFormat

Case "Blank"

GetCellSingleValue = IsCellBlank(cell)

Case "Error"

GetCellSingleValue = IsError(cell.Value)

End Select

End Function

'==================================================
' Excel UDF Lib Ver2.0
' Module : modCell
' Part3 : Last Row / Last Column
'==================================================

'==================================================
' CellLastRow
'==================================================

Public Function CellLastRow(Optional ByVal Target As Variant) As Variant

Dim cell As Range

If IsHelp(Target) Then

CellLastRow = HelpCellLastRow()

Exit Function

End If

If IsMissing(Target) Then

Set cell = Application.Caller

Else

Set cell = GetTargetRange(Target)

End If

If cell Is Nothing Then

CellLastRow = ReturnValueError()

Exit Function

End If

CellLastRow = GetLastRow(cell)

End Function

'==================================================
' CellLastColumn
'==================================================

Public Function CellLastColumn(Optional ByVal Target As Variant) As Variant

Dim cell As Range

If IsHelp(Target) Then

CellLastColumn = HelpCellLastColumn()

Exit Function

End If

If IsMissing(Target) Then

Set cell = Application.Caller

Else

Set cell = GetTargetRange(Target)

End If

If cell Is Nothing Then

CellLastColumn = ReturnValueError()

Exit Function

End If

CellLastColumn = GetLastColumn(cell)

End Function

'==================================================
' GetLastRow
'==================================================

Private Function GetLastRow(ByVal Target As Range) As Long

Dim ws As Worksheet
Dim col As Long

Set ws = Target.Worksheet

col = Target.Column

GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

End Function

'==================================================
' GetLastColumn
'==================================================

Private Function GetLastColumn(ByVal Target As Range) As Long

Dim ws As Worksheet
Dim Row As Long

Set ws = Target.Worksheet

Row = Target.Row

GetLastColumn = ws.Cells(Row, ws.Columns.Count).End(xlToLeft).Column

End Function

'==================================================
' Excel UDF Lib Ver2.0
' Module : modCell
' Part4 : Help
'==================================================

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

Next i

HelpArray = Result

End Function

'==================================================
' CellInfo Help
'==================================================

Private Function HelpCellInfo() As Variant

HelpCellInfo = HelpArray("Function : CellInfo" & vbLf & "Description : Returns cell information." & vbLf & "Syntax : =CellInfo(Target)" & vbLf & "" & vbLf & "Input :" & vbLf & "Cell reference supported." & vbLf & "UDF result supported." & vbLf & "Multiple cells range supported." & vbLf & "Blank input returns caller cell." & vbLf & "" & vbLf & "Output :" & vbLf & "Dynamic Array." & vbLf & "Header included." & vbLf & "Vertical spill." & vbLf & "Column priority." & vbLf & "" & vbLf & "Version : Ver2.0")

End Function

'==================================================
' Single Function Help
'==================================================

Private Function HelpSingle(ByVal FunctionName As String, ByVal Description As String, ByVal Syntax As String, ByVal Output As String) As Variant

HelpSingle = HelpArray("Function : " & FunctionName & vbLf & "Description : " & Description & vbLf & "Syntax : " & Syntax & vbLf & "" & vbLf & "Input :" & vbLf & "Single cell only." & vbLf & "Cell reference supported." & vbLf & "UDF result supported." & vbLf & "" & vbLf & "Output :" & vbLf & Output & vbLf & "" & vbLf & "Version : Ver2.0")

End Function

'==================================================
' CellType Help
'==================================================

Private Function HelpCellType() As Variant

HelpCellType = HelpSingle("CellType", "Returns cell type.", "=CellType(Target)", "Cell type.")

End Function

'==================================================
' CellHasFormula Help
'==================================================

Private Function HelpCellHasFormula() As Variant

HelpCellHasFormula = HelpSingle("CellHasFormula", "Checks formula existence.", "=CellHasFormula(Target)", "TRUE / FALSE.")

End Function

'==================================================
' CellFormula Help
'==================================================

Private Function HelpCellFormula() As Variant

HelpCellFormula = HelpSingle("CellFormula", "Returns cell formula.", "=CellFormula(Target)", "Formula text. Blank if no formula.")

End Function

'==================================================
' CellFormat Help
'==================================================

Private Function HelpCellFormat() As Variant

HelpCellFormat = HelpSingle("CellFormat", "Returns cell number format.", "=CellFormat(Target)", "NumberFormat.")

End Function

'==================================================
' CellBlank Help
'==================================================

Private Function HelpCellBlank() As Variant

HelpCellBlank = HelpSingle("CellBlank", "Checks blank cell.", "=CellBlank(Target)", "TRUE / FALSE.")

End Function

'==================================================
' CellError Help
'==================================================

Private Function HelpCellError() As Variant

HelpCellError = HelpSingle("CellError", "Checks error cell.", "=CellError(Target)", "TRUE / FALSE.")

End Function

'==================================================
' CellLastRow Help
'==================================================

Private Function HelpCellLastRow() As Variant

HelpCellLastRow = HelpSingle("CellLastRow", "Returns last row in column.", "=CellLastRow(Target)", "Last used row number.")

End Function

'==================================================
' CellLastColumn Help
'==================================================

Private Function HelpCellLastColumn() As Variant

HelpCellLastColumn = HelpSingle("CellLastColumn", "Returns last column in row.", "=CellLastColumn(Target)", "Last used column number.")

End Function

'==================================================
' Excel UDF Lib Ver2.0
' Module : modCell
' Part5 : Specification / Final Check
'==================================================

' (略: 仕様コメントブロックは元のまま)
'==================================================
