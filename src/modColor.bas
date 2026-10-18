Attribute VB_Name = "ModColor"
Option Explicit

'==================================================
' Excel UDF Lib Ver2.0
' Module : modColor
' Color Information UDF Library
' Compact Final (修正版)
' 修正点: Target 解決を呼び出し元シート優先に変更（アドイン自身のシート参照回避）
'==================================================

Private Const HELP_MARK As String = "?"

'========================
' 共通ユーティリティ
'========================

Private Function ReturnNA() As Variant
    ReturnNA = CVErr(xlErrNA)
End Function

Private Function GetText(ByVal Value As Variant) As String
    If IsObject(Value) Then Exit Function
    If IsArray(Value) Then Exit Function
    If IsError(Value) Then Exit Function
    If IsEmpty(Value) Then Exit Function
    GetText = CStr(Value)
End Function

Private Function IsHelp(ByVal Target As Variant) As Boolean
    If IsObject(Target) Then Exit Function
    IsHelp = (GetText(Target) = HELP_MARK)
End Function

' 呼び出し元ワークシート（セルから呼ばれた場合はそのセルの属する Worksheet）を返す
Private Function GetCallerWorksheet() As Worksheet
    On Error Resume Next
    If TypeName(Application.Caller) = "Range" Then
        Set GetCallerWorksheet = Application.Caller.Parent
    Else
        Set GetCallerWorksheet = Application.ActiveSheet
    End If
    On Error GoTo 0
End Function

' Target が Range のときはそのまま返す。
' 文字列の場合はまず呼び出し元シートで解釈を試み、それでダメなら Application.Range を試す。
Private Function ResolveRange(ByVal Target As Variant) As Range
    Dim ws As Worksheet
    Dim rng As Range
    Dim nm As Name

    On Error Resume Next
    If TypeName(Target) = "Range" Then
        Set ResolveRange = Target
        On Error GoTo 0
        Exit Function
    End If

    ' 空・非文字列は Nothing
    If IsMissing(Target) Or IsError(Target) Then
        On Error GoTo 0
        Exit Function
    End If

    If Not IsObject(Target) Then
        ws = GetCallerWorksheet()
        If Not ws Is Nothing Then
            Set rng = Nothing
            Set rng = ws.Range(CStr(Target))
            If Not rng Is Nothing Then
                Set ResolveRange = rng
                On Error GoTo 0
                Exit Function
            End If
        End If

        ' 呼び出し元シートで解決できなければ Application.Range（外部参照等）を試す
        Set rng = Nothing
        Set rng = Application.Range(CStr(Target))
        If Not rng Is Nothing Then
            Set ResolveRange = rng
            On Error GoTo 0
            Exit Function
        End If

        ' 名前定義（ThisWorkbook 内の名前）も最後に試す（補助）
        For Each nm In ThisWorkbook.Names
            If LCase(nm.Name) = LCase(CStr(Target)) Then
                On Error Resume Next
                Set ResolveRange = nm.RefersToRange
                On Error GoTo 0
                If Not ResolveRange Is Nothing Then Exit Function
            End If
        Next nm
    End If
    On Error GoTo 0
End Function

'========================
' Help ユーティリティ
'========================

Private Function HelpArray( _
    ByVal FunctionName As String, _
    ByVal Description As String, _
    ByVal Syntax As String _
) As Variant
    Dim arr(1 To 4, 1 To 2) As Variant
    arr(1, 1) = "Function : " & FunctionName
    arr(2, 1) = "Description : " & Description
    arr(3, 1) = "Syntax : " & Syntax
    arr(4, 1) = "Version : " & "Excel UDF Lib Ver2.0"
    arr(1, 2) = ""
    arr(2, 2) = ""
    arr(3, 2) = ""
    arr(4, 2) = ""
    HelpArray = arr
End Function

Private Function HelpSingle( _
    ByVal FunctionName As String, _
    ByVal Description As String, _
    ByVal Syntax As String, _
    ByVal Output As String _
) As Variant
    Dim Text As String
    Dim Lines() As String
    Dim arr() As Variant
    Dim i As Long

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
    "Version : Excel UDF Lib Ver2.0"

    Lines = Split(Text, vbLf)
    ReDim arr(1 To UBound(Lines) + 1, 1 To 1)

    For i = 0 To UBound(Lines)
        arr(i + 1, 1) = Lines(i)
    Next i

    HelpSingle = arr
End Function

'========================
' Color Helpers
'========================

Private Function GetRGB(ByVal Color As Long) As String
    GetRGB = _
    (Color And &HFF&) & "," & _
    ((Color \ &H100&) And &HFF&) & "," & _
    ((Color \ &H10000) And &HFF&)
End Function

Private Function GetHex(ByVal Color As Long) As String
    GetHex = _
    "#" & _
    Right$("0" & Hex$(Color And &HFF&), 2) & _
    Right$("0" & Hex$((Color \ &H100&) And &HFF&), 2) & _
    Right$("0" & Hex$((Color \ &H10000) And &HFF&), 2)
End Function

'========================
' ColorInfo 実装
'========================

Private Function ColorInfoSingle( _
    ByVal Target As Range _
) As Variant
    Dim arr(1 To 2, 1 To 6) As Variant
    Dim cellColor As Long
    Dim fontColor As Long

    arr(1, 1) = "Address"
    arr(1, 2) = "Value"
    arr(1, 3) = "CellRGB"
    arr(1, 4) = "CellHex"
    arr(1, 5) = "FontRGB"
    arr(1, 6) = "FontHex"

    arr(2, 1) = Target.Address(False, False)

    If IsEmpty(Target.Value) Then
        arr(2, 2) = ""
    Else
        arr(2, 2) = Target.Value
    End If

    cellColor = Target.Interior.Color
    fontColor = Target.Font.Color

    arr(2, 3) = GetRGB(cellColor)
    arr(2, 4) = GetHex(cellColor)
    arr(2, 5) = GetRGB(fontColor)
    arr(2, 6) = GetHex(fontColor)

    ColorInfoSingle = arr
End Function

Private Function ColorInfoMulti( _
    ByVal Target As Range _
) As Variant
    Dim arr() As Variant
    Dim cell As Range
    Dim Row As Long
    Dim temp As Variant
    Dim col As Long

    ReDim arr(1 To Target.Cells.Count + 1, 1 To 6)

    arr(1, 1) = "Address"
    arr(1, 2) = "Value"
    arr(1, 3) = "CellRGB"
    arr(1, 4) = "CellHex"
    arr(1, 5) = "FontRGB"
    arr(1, 6) = "FontHex"

    Row = 2

    For Each cell In Target.Cells
        temp = ColorInfoSingle(cell)
        For col = 1 To 6
            arr(Row, col) = temp(2, col)
        Next col
        Row = Row + 1
    Next cell

    ColorInfoMulti = arr
End Function

Public Function ColorInfo( _
    Optional ByVal Target As Variant _
) As Variant

    If IsHelp(Target) Then
        ColorInfo = HelpSingle( _
            "ColorInfo", _
            "Returns cell color information.", _
            "=ColorInfo(Target)", _
            "Address,Value,CellRGB,CellHex,FontRGB,FontHex")
        Exit Function
    End If

    Dim rng As Range

    If IsMissing(Target) Then
        ' 省略時は呼び出し元セルを試す（アドインから呼ばれても呼び出し元を優先）
        On Error Resume Next
        If TypeName(Application.Caller) = "Range" Then
            Set rng = Application.Caller
        End If
        On Error GoTo 0
        If rng Is Nothing Then
            ColorInfo = ReturnNA()
            Exit Function
        End If
    Else
        Set rng = ResolveRange(Target)
        If rng Is Nothing Then
            ColorInfo = ReturnNA()
            Exit Function
        End If
    End If

    If rng.Cells.Count > 1 Then
        ColorInfo = ColorInfoMulti(rng)
    Else
        ColorInfo = ColorInfoSingle(rng)
    End If
End Function

'========================
' 個別関数群（Hex / RGB 取得）
'========================

Public Function CellColorHex( _
    ByVal Target As Variant _
) As Variant

    If IsHelp(Target) Then
        CellColorHex = HelpSingle( _
            "CellColorHex", _
            "Returns cell fill color in Hex.", _
            "=CellColorHex(Target)", _
            "Hex color")
        Exit Function
    End If

    Dim rng As Range
    Set rng = ResolveRange(Target)
    If rng Is Nothing Then
        CellColorHex = ReturnNA()
        Exit Function
    End If

    CellColorHex = GetHex(rng.Interior.Color)
End Function

Public Function CellColorRGB( _
    ByVal Target As Variant _
) As Variant

    If IsHelp(Target) Then
        CellColorRGB = HelpSingle( _
            "CellColorRGB", _
            "Returns cell fill color in RGB.", _
            "=CellColorRGB(Target)", _
            "R,G,B")
        Exit Function
    End If

    Dim rng2 As Range
    Set rng2 = ResolveRange(Target)
    If rng2 Is Nothing Then
        CellColorRGB = ReturnNA()
        Exit Function
    End If

    CellColorRGB = GetRGB(rng2.Interior.Color)
End Function

Public Function FontColorHex( _
    ByVal Target As Variant _
) As Variant

    If IsHelp(Target) Then
        FontColorHex = HelpSingle( _
            "FontColorHex", _
            "Returns font color in Hex.", _
            "=FontColorHex(Target)", _
            "Hex color")
        Exit Function
    End If

    Dim rng3 As Range
    Set rng3 = ResolveRange(Target)
    If rng3 Is Nothing Then
        FontColorHex = ReturnNA()
        Exit Function
    End If

    FontColorHex = GetHex(rng3.Font.Color)
End Function

Public Function FontColorRGB( _
    ByVal Target As Variant _
) As Variant

    If IsHelp(Target) Then
        FontColorRGB = HelpSingle( _
            "FontColorRGB", _
            "Returns font color in RGB.", _
            "=FontColorRGB(Target)", _
            "R,G,B")
        Exit Function
    End If

    Dim rng4 As Range
    Set rng4 = ResolveRange(Target)
    If rng4 Is Nothing Then
        FontColorRGB = ReturnNA()
        Exit Function
    End If

    FontColorRGB = GetRGB(rng4.Font.Color)
End Function


