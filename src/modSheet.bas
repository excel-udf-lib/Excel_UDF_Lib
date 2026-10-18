Attribute VB_Name = "ModSheet"
Option Explicit
'==================================================
' Excel UDF Lib Ver2.0 (修正版)
' Module : modSheet
' 修正点: ThisWorkbook の参照を呼び出し元ブックに置換
'==================================================

'------------------------
' 共通ヘルパー
'------------------------
' 呼び出し元の Workbook を返す（セルから呼ばれた場合はそのセルの属する Workbook）
Private Function GetCallerWorkbook() As Workbook
    On Error Resume Next
    If TypeName(Application.Caller) = "Range" Then
        Set GetCallerWorkbook = Application.Caller.Parent.Parent
    Else
        ' セル外から呼び出された場合の保険
        Set GetCallerWorkbook = Application.ActiveWorkbook
    End If
    On Error GoTo 0
End Function

' 指定されたコンテキスト（引数）に基づいてワークシートを探す
' lookupTarget が省略または空のときは GetCallerWorkbook を使う
Private Function FindWorksheetByName(ByVal SheetKey As String, Optional ByVal lookupTarget As Variant) As Worksheet
    Dim wb As Workbook
    Dim w As Workbook
    On Error Resume Next

    If IsMissing(lookupTarget) Then
        Set wb = GetCallerWorkbook()
    ElseIf TypeName(lookupTarget) = "Range" Then
        ' 空セルだったら呼び出し元ブックを使う
        If lookupTarget.Cells.Count = 1 Then
            If IsEmpty(lookupTarget.Value) Then
                Set wb = GetCallerWorkbook()
            Else
                Set wb = lookupTarget.Worksheet.Parent
            End If
        Else
            Set wb = GetCallerWorkbook()
        End If
    ElseIf TypeName(lookupTarget) = "String" Then
        For Each w In Application.Workbooks
            If LCase(w.Name) = LCase(CStr(lookupTarget)) Or LCase(w.FullName) = LCase(CStr(lookupTarget)) Then
                Set wb = w
                Exit For
            End If
        Next w
    Else
        Set wb = GetCallerWorkbook()
    End If

    If Not wb Is Nothing Then
        On Error Resume Next
        Set FindWorksheetByName = wb.Worksheets(SheetKey)
        On Error GoTo 0
    End If
End Function

'==================================================
' SheetExists
'==================================================
Public Function SheetExists(Optional ByVal Target As Variant) As Variant
    Dim ws As Worksheet
    Dim SheetKey As String

    If Not IsMissing(Target) Then
        If GetInputText(Target) = "?" Then
            SheetExists = HelpSheetExists()
            Exit Function
        End If
    End If

    If IsMissing(Target) Then
        SheetExists = True
        Exit Function
    End If

    If TypeName(Target) = "Range" Then
        If Target.Cells.Count <> 1 Then
            SheetExists = CVErr(xlErrValue)
            Exit Function
        End If

        If IsEmpty(Target.Value) Then
            SheetExists = True
            Exit Function
        End If
    End If

    SheetKey = GetInputText(Target)

    Set ws = FindWorksheetByName(SheetKey, Target)

    SheetExists = Not ws Is Nothing
End Function

'==================================================
' SheetIndex
'==================================================
Public Function SheetIndex(Optional ByVal Target As Variant) As Variant
    Dim ws As Worksheet
    Dim SheetKey As String

    If IsMissing(Target) Then
        On Error Resume Next
        SheetIndex = Application.Caller.Worksheet.index
        On Error GoTo 0
        Exit Function
    End If

    If GetInputText(Target) = "?" Then
        SheetIndex = HelpSheetIndex()
        Exit Function
    End If

    If TypeName(Target) = "Range" Then
        If Target.Cells.Count <> 1 Then
            SheetIndex = CVErr(xlErrValue)
            Exit Function
        End If

        If IsEmpty(Target.Value) Then
            SheetIndex = Target.Worksheet.index
            Exit Function
        End If
    End If

    SheetKey = GetInputText(Target)

    Set ws = FindWorksheetByName(SheetKey, Target)

    If ws Is Nothing Then
        SheetIndex = CVErr(xlErrNA)
    Else
        SheetIndex = ws.index
    End If
End Function

'==================================================
' SheetName
'==================================================
Public Function SheetName(Optional ByVal Target As Variant) As Variant
    Dim ws As Worksheet
    Dim SheetKey As String

    If IsMissing(Target) Then
        On Error Resume Next
        SheetName = Application.Caller.Worksheet.Name
        On Error GoTo 0
        Exit Function
    End If

    If GetInputText(Target) = "?" Then
        SheetName = HelpSheetName()
        Exit Function
    End If

    If TypeName(Target) = "Range" Then
        If Target.Cells.Count <> 1 Then
            SheetName = CVErr(xlErrValue)
            Exit Function
        End If

        If IsEmpty(Target.Value) Then
            SheetName = Target.Worksheet.Name
            Exit Function
        End If
    End If

    SheetKey = GetInputText(Target)

    Set ws = FindWorksheetByName(SheetKey, Target)

    If ws Is Nothing Then
        SheetName = CVErr(xlErrNA)
    Else
        SheetName = ws.Name
    End If
End Function

'==================================================
' SheetCount
'==================================================
Public Function SheetCount(Optional ByVal Target As Variant) As Variant
    Dim wb As Workbook

    If Not IsMissing(Target) Then
        If GetInputText(Target) = "?" Then
            SheetCount = HelpSheetCount()
            Exit Function
        Else
            SheetCount = CVErr(xlErrValue)
            Exit Function
        End If
    End If

    Set wb = GetCallerWorkbook()
    If wb Is Nothing Then
        SheetCount = CVErr(xlErrRef)
    Else
        SheetCount = wb.Worksheets.Count
    End If
End Function

'==================================================
' SheetInfo
'==================================================
Public Function SheetInfo(Optional ByVal Target As Variant) As Variant
    Dim inputs As Collection
    Dim ws As Worksheet
    Dim Result() As Variant
    Dim i As Long
    Dim rowCount As Long
    Dim wb As Workbook

    If IsMissing(Target) Then
        Set inputs = New Collection
        On Error Resume Next
        inputs.Add Application.Caller.Worksheet.Name
        On Error GoTo 0
    ElseIf GetInputText(Target) = "?" Then
        SheetInfo = HelpSheetInfo()
        Exit Function
    Else
        Set inputs = GetSheetInputs(Target)
    End If

    rowCount = inputs.Count
    ReDim Result(1 To rowCount + 1, 1 To 3)

    Result(1, 1) = "SheetName"
    Result(1, 2) = "SheetIndex"
    Result(1, 3) = "SheetExists"

    For i = 1 To rowCount
        Set ws = Nothing
        Set ws = FindWorksheetByName(CStr(inputs(i)), Target)

        If ws Is Nothing Then
            Result(i + 1, 1) = CVErr(xlErrNA)
            Result(i + 1, 2) = CVErr(xlErrNA)
            Result(i + 1, 3) = False
        Else
            Result(i + 1, 1) = ws.Name
            Result(i + 1, 2) = ws.index
            Result(i + 1, 3) = True
        End If
    Next i

    SheetInfo = Result
End Function

'==================================================
' Sheet Inputs
'==================================================
Private Function GetSheetInputs(ByVal Target As Variant) As Collection
    Dim Result As New Collection
    Dim c As Range

    If TypeName(Target) = "Range" Then
        For Each c In Target.Cells
            If Not IsEmpty(c.Value) Then
                Result.Add CStr(c.Value)
            End If
        Next c
    Else
        If CStr(Target) <> "" Then
            Result.Add CStr(Target)
        End If
    End If

    Set GetSheetInputs = Result
End Function

'==================================================
' SheetList
'==================================================
Public Function SheetList(Optional ByVal Target As Variant) As Variant
    Dim ws As Worksheet
    Dim Result() As Variant
    Dim i As Long
    Dim Count As Long
    Dim wb As Workbook

    If Not IsMissing(Target) Then
        If GetInputText(Target) = "?" Then
            SheetList = HelpSheetList()
        Else
            SheetList = CVErr(xlErrValue)
        End If
        Exit Function
    End If

    Set wb = GetCallerWorkbook()
    If wb Is Nothing Then
        SheetList = CVErr(xlErrRef)
        Exit Function
    End If

    Count = wb.Worksheets.Count

    ReDim Result(1 To Count + 1, 1 To 3)

    Result(1, 1) = "SheetName"
    Result(1, 2) = "SheetIndex"
    Result(1, 3) = "SheetVisible"

    For i = 1 To Count
        Set ws = wb.Worksheets(i)
        Result(i + 1, 1) = ws.Name
        Result(i + 1, 2) = ws.index
        Result(i + 1, 3) = GetVisibleText(ws.Visible)
    Next i

    SheetList = Result
End Function

'==================================================
' Visible Text
'==================================================
Private Function GetVisibleText(ByVal Value As XlSheetVisibility) As String
    Select Case Value
        Case xlSheetVisible
            GetVisibleText = "Visible"
        Case xlSheetHidden
            GetVisibleText = "Hidden"
        Case xlSheetVeryHidden
            GetVisibleText = "VeryHidden"
    End Select
End Function

'==================================================
' Input Text
'==================================================
Private Function GetInputText(ByVal Target As Variant) As String
    If TypeName(Target) = "Range" Then
        GetInputText = CStr(Target.Cells(1, 1).Value)
    Else
        GetInputText = CStr(Target)
    End If
End Function

'==================================================
' Help : SheetExists
'==================================================
Private Function HelpSheetExists() As Variant
    Dim Result(1 To 12, 1 To 1) As Variant

    Result(1, 1) = "Function : SheetExists"
    Result(2, 1) = "Description : Checks sheet existence."
    Result(3, 1) = "Syntax : =SheetExists(SheetName)"
    Result(4, 1) = ""
    Result(5, 1) = "Input :"
    Result(6, 1) = "Cell reference supported."
    Result(7, 1) = "UDF result supported."
    Result(8, 1) = ""
    Result(9, 1) = "Output :"
    Result(10, 1) = "TRUE or FALSE."
    Result(11, 1) = ""
    Result(12, 1) = "Version : Ver2.0"

    HelpSheetExists = Result
End Function

'==================================================
' Help : SheetIndex
'==================================================
Private Function HelpSheetIndex() As Variant
    Dim Result(1 To 12, 1 To 1) As Variant

    Result(1, 1) = "Function : SheetIndex"
    Result(2, 1) = "Description : Returns sheet index."
    Result(3, 1) = "Syntax : =SheetIndex(SheetName)"
    Result(4, 1) = ""
    Result(5, 1) = "Input :"
    Result(6, 1) = "Cell reference supported."
    Result(7, 1) = "UDF result supported."
    Result(8, 1) = "Blank input returns current sheet."
    Result(9, 1) = ""
    Result(10, 1) = "Output :"
    Result(11, 1) = "Sheet index."
    Result(12, 1) = "Version : Ver2.0"

    HelpSheetIndex = Result
End Function

'==================================================
' Help : SheetName
'==================================================
Private Function HelpSheetName() As Variant
    Dim Result(1 To 12, 1 To 1) As Variant

    Result(1, 1) = "Function : SheetName"
    Result(2, 1) = "Description : Returns sheet name."
    Result(3, 1) = "Syntax : =SheetName(SheetName)"
    Result(4, 1) = ""
    Result(5, 1) = "Input :"
    Result(6, 1) = "Cell reference supported."
    Result(7, 1) = "UDF result supported."
    Result(8, 1) = "Blank input returns current sheet."
    Result(9, 1) = ""
    Result(10, 1) = "Output :"
    Result(11, 1) = "Sheet name."
    Result(12, 1) = "Version : Ver2.0"

    HelpSheetName = Result
End Function

'==================================================
' Help : SheetCount
'==================================================
Private Function HelpSheetCount() As Variant
    Dim Result(1 To 12, 1 To 1) As Variant

    Result(1, 1) = "Function : SheetCount"
    Result(2, 1) = "Description : Returns sheet count."
    Result(3, 1) = "Syntax : =SheetCount()"
    Result(4, 1) = ""
    Result(5, 1) = "Input :"
    Result(6, 1) = "No argument required."
    Result(7, 1) = "Use ? for help."
    Result(8, 1) = ""
    Result(9, 1) = "Output :"
    Result(10, 1) = "Sheet count."
    Result(11, 1) = ""
    Result(12, 1) = "Version : Ver2.0"

    HelpSheetCount = Result
End Function

'==================================================
' Help : SheetInfo
'==================================================
Private Function HelpSheetInfo() As Variant
    Dim Result(1 To 16, 1 To 1) As Variant

    Result(1, 1) = "Function : SheetInfo"
    Result(2, 1) = "Description : Returns sheet information."
    Result(3, 1) = "Syntax : =SheetInfo(SheetName)"
    Result(4, 1) = ""
    Result(5, 1) = "Input :"
    Result(6, 1) = "Cell reference supported."
    Result(7, 1) = "UDF result supported."
    Result(8, 1) = "Multiple cells range supported."
    Result(9, 1) = "Blank cells ignored."
    Result(10, 1) = ""
    Result(11, 1) = "Output :"
    Result(12, 1) = "Dynamic Array."
    Result(13, 1) = "Header included."
    Result(14, 1) = "Vertical spill."
    Result(15, 1) = ""
    Result(16, 1) = "Version : Ver2.0"

    HelpSheetInfo = Result
End Function

'==================================================
' Help : SheetList
'==================================================
Private Function HelpSheetList() As Variant
    Dim Result(1 To 14, 1 To 1) As Variant

    Result(1, 1) = "Function : SheetList"
    Result(2, 1) = "Description : Returns sheet list."
    Result(3, 1) = "Syntax : =SheetList()"
    Result(4, 1) = ""
    Result(5, 1) = "Input :"
    Result(6, 1) = "No argument required."
    Result(7, 1) = "Use ? for help."
    Result(8, 1) = ""
    Result(9, 1) = "Output :"
    Result(10, 1) = "Dynamic Array."
    Result(11, 1) = "Header included."
    Result(12, 1) = "Vertical spill."
    Result(13, 1) = ""
    Result(14, 1) = "Version : Ver2.0"

    HelpSheetList = Result
End Function

