Attribute VB_Name = "modText"
' Excel UDF Lib Ver2.0.0 Fix
' Module : modText (修正版)
' Text Processing UDF Library
Option Explicit

Private Const HELP_MARK As String = "?"

'==================================================
' Common utilities: caller-sheet 優先で文字列/Range を解決
'==================================================

' 呼び出し元の Worksheet を返す（セルから呼ばれた場合はそのセルの属する Worksheet）
Private Function GetCallerWorksheet() As Worksheet
    On Error Resume Next
    If TypeName(Application.Caller) = "Range" Then
        Set GetCallerWorksheet = Application.Caller.Parent
    Else
        Set GetCallerWorksheet = Application.ActiveSheet
    End If
    On Error GoTo 0
End Function

' 文字列 Target をまず呼び出し元シートで Range として解釈し、解決できなければ Application.Range を試す
' 解決できれば Range を返す。解決できない/無効な入力なら Nothing を返す。
Private Function ResolveRange(ByVal Target As Variant) As Range
    Dim ws As Worksheet
    Dim rng As Range
    On Error Resume Next

    If TypeName(Target) = "Range" Then
        Set ResolveRange = Target
        On Error GoTo 0
        Exit Function
    End If

    If IsMissing(Target) Or IsError(Target) Then
        On Error GoTo 0
        Exit Function
    End If

    If Not IsObject(Target) Then
        Set ws = GetCallerWorksheet()
        If Not ws Is Nothing Then
            Set rng = Nothing
            Set rng = ws.Range(CStr(Target))
            If Not rng Is Nothing Then
                Set ResolveRange = rng
                On Error GoTo 0
                Exit Function
            End If
        End If

        Set rng = Nothing
        Set rng = Application.Range(CStr(Target))
        If Not rng Is Nothing Then
            Set ResolveRange = rng
            On Error GoTo 0
            Exit Function
        End If
    End If

    On Error GoTo 0
End Function

' GetInputText: Target（Range か文字列）からテキストを取り出す（呼び出し元優先）
Private Function GetInputText(ByVal Target As Variant) As String
    Dim rng As Range
    On Error Resume Next

    If TypeName(Target) = "Range" Then
        GetInputText = CStr(Target.Cells(1, 1).Value)
        On Error GoTo 0
        Exit Function
    End If

    ' 文字列やその他が渡された場合、まずそれをセル参照として解決してみる
    Set rng = ResolveRange(Target)
    If Not rng Is Nothing Then
        GetInputText = CStr(rng.Cells(1, 1).Value)
        On Error GoTo 0
        Exit Function
    End If

    ' それ以外は単純に文字列化
    GetInputText = CStr(Target)
    On Error GoTo 0
End Function

Private Function IsHelp(ByVal Target As Variant) As Boolean
    IsHelp = (GetInputText(Target) = HELP_MARK)
End Function

Private Function ReturnNA() As Variant
    ReturnNA = CVErr(xlErrNA)
End Function

Private Function ReturnValueError() As Variant
    ReturnValueError = CVErr(xlErrValue)
End Function

'==================================================
' Input helper for single-text extraction (呼び出し元優先)
'  - 戻り値 True = Text に取り出せた、False = 引数不正
'==================================================
Private Function GetTextInput( _
    ByVal Target As Variant, _
    ByRef Text As String _
) As Boolean

    Dim rng As Range

    If TypeName(Target) = "Range" Then
        If Target.Cells.Count <> 1 Then Exit Function
        Text = CStr(Target.Value)
        GetTextInput = True
        Exit Function
    End If

    ' 文字列等の場合、まずそれを範囲参照として解決して値を取れるか試す
    Set rng = ResolveRange(Target)
    If Not rng Is Nothing Then
        If rng.Cells.Count <> 1 Then Exit Function
        Text = CStr(rng.Value)
        GetTextInput = True
        Exit Function
    End If

    ' 最後に、単なる文字列として受け取る
    On Error Resume Next
    Text = CStr(Target)
    If Err.Number = 0 Then
        GetTextInput = True
    Else
        GetTextInput = False
    End If
    On Error GoTo 0
End Function

'==================================================
' Number utilities
'==================================================

Private Function IsNumericCharacter( _
    ByVal Character As String _
) As Boolean

    Dim Code As Long

    If Character = "" Then Exit Function

    Code = AscW(Character)
    If Code < 0 Then Code = Code + 65536

    IsNumericCharacter = _
        ((Code >= 48 And Code <= 57) Or _
         (Code >= 65296 And Code <= 65305))

End Function

Private Function NormalizeNumericCharacter( _
    ByVal Character As String _
) As String

    Dim Code As Long

    Code = AscW(Character)
    If Code < 0 Then Code = Code + 65536

    If Code >= 65296 And Code <= 65305 Then
        NormalizeNumericCharacter = ChrW(Code - 65248)
    Else
        NormalizeNumericCharacter = Character
    End If

End Function

'==================================================
' NumOnly / NumSpill 実装
'==================================================

Private Function ExtractNumberText( _
    ByVal Text As String _
) As Variant

    Dim Result As String
    Dim i As Long
    Dim Character As String

    For i = 1 To Len(Text)
        Character = Mid$(Text, i, 1)
        If IsNumericCharacter(Character) Then
            Result = Result & NormalizeNumericCharacter(Character)
        End If
    Next i

    If Result = "" Then
        ExtractNumberText = ReturnNA()
    Else
        ExtractNumberText = Result
    End If
End Function

Private Function ExtractNumberSpill( _
    ByVal Text As String _
) As Variant

    Dim Items As Collection
    Dim Result() As Variant
    Dim CurrentText As String
    Dim Character As String
    Dim i As Long
    Dim j As Long

    Set Items = New Collection

    For i = 1 To Len(Text)
        Character = Mid$(Text, i, 1)
        If IsNumericCharacter(Character) Then
            CurrentText = CurrentText & NormalizeNumericCharacter(Character)
        Else
            If CurrentText <> "" Then
                Items.Add CurrentText
                CurrentText = ""
            End If
        End If
    Next i

    If CurrentText <> "" Then Items.Add CurrentText

    If Items.Count = 0 Then
        ExtractNumberSpill = ReturnNA()
        Exit Function
    End If

    ReDim Result(1 To Items.Count, 1 To 1)
    For j = 1 To Items.Count
        Result(j, 1) = Items(j)
    Next j

    ExtractNumberSpill = Result
End Function

Public Function NumOnly( _
    ByVal Target As Variant _
) As Variant

    Dim Text As String

    If IsHelp(Target) Then
        NumOnly = HelpNumOnly()
        Exit Function
    End If

    If Not GetTextInput(Target, Text) Then
        NumOnly = ReturnValueError()
        Exit Function
    End If

    NumOnly = ExtractNumberText(Text)
End Function

Public Function NumSpill( _
    ByVal Target As Variant _
) As Variant

    Dim Text As String

    If IsHelp(Target) Then
        NumSpill = HelpNumSpill()
        Exit Function
    End If

    If Not GetTextInput(Target, Text) Then
        NumSpill = ReturnValueError()
        Exit Function
    End If

    NumSpill = ExtractNumberSpill(Text)
End Function

'==================================================
' StrInfo
'==================================================

Public Function StrInfo( _
    ByVal Target As Variant _
) As Variant

    Dim Text As String

    If IsHelp(Target) Then
        StrInfo = HelpStrInfo()
        Exit Function
    End If

    If Not GetTextInput(Target, Text) Then
        StrInfo = ReturnValueError()
        Exit Function
    End If

    StrInfo = BuildStrInfo(Text)
End Function

Private Function BuildStrInfo( _
    ByVal Text As String _
) As Variant

    Dim Work() As Variant
    Dim Result() As Variant
    Dim i As Long
    Dim Row As Long
    Dim StartPos As Long
    Dim BlockLength As Long
    Dim BlockText As String
    Dim Character As String
    Dim CurrentNumber As Boolean
    Dim NextNumber As Boolean

    ReDim Work(1 To Len(Text) + 1, 1 To 4)

    Work(1, 1) = "Block"
    Work(1, 2) = "Start"
    Work(1, 3) = "Length"
    Work(1, 4) = "Type"

    If Len(Text) = 0 Then
        BuildStrInfo = Work
        Exit Function
    End If

    StartPos = 1
    BlockText = Mid$(Text, 1, 1)
    BlockLength = 1
    CurrentNumber = IsNumericCharacter(BlockText)

    For i = 2 To Len(Text)
        Character = Mid$(Text, i, 1)
        NextNumber = IsNumericCharacter(Character)
        If NextNumber = CurrentNumber Then
            BlockText = BlockText & Character
            BlockLength = BlockLength + 1
        Else
            Row = Row + 1
            Work(Row + 1, 1) = NormalizeBlock(BlockText, CurrentNumber)
            Work(Row + 1, 2) = StartPos
            Work(Row + 1, 3) = BlockLength
            Work(Row + 1, 4) = IIf(CurrentNumber, "Number", "Text")
            StartPos = i
            BlockText = Character
            BlockLength = 1
            CurrentNumber = NextNumber
        End If
    Next i

    Row = Row + 1
    Work(Row + 1, 1) = NormalizeBlock(BlockText, CurrentNumber)
    Work(Row + 1, 2) = StartPos
    Work(Row + 1, 3) = BlockLength
    Work(Row + 1, 4) = IIf(CurrentNumber, "Number", "Text")

    ReDim Result(1 To Row + 1, 1 To 4)
    For i = 1 To Row + 1
        Result(i, 1) = Work(i, 1)
        Result(i, 2) = Work(i, 2)
        Result(i, 3) = Work(i, 3)
        Result(i, 4) = Work(i, 4)
    Next i

    BuildStrInfo = Result
End Function

Private Function NormalizeBlock( _
    ByVal Text As String, _
    ByVal IsNumber As Boolean _
) As String

    Dim i As Long

    If Not IsNumber Then
        NormalizeBlock = Text
        Exit Function
    End If

    For i = 1 To Len(Text)
        NormalizeBlock = NormalizeBlock & NormalizeNumericCharacter(Mid$(Text, i, 1))
    Next i
End Function

'==================================================
' Help テキスト（変更なし）
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

Private Function HelpNumOnly() As Variant
    Dim Text As String
    Text = "Function : NumOnly" & vbLf
    Text = Text & "Description : Extracts numeric characters from text and returns them as a text string." & vbLf
    Text = Text & "Syntax : =NumOnly(Target)" & vbLf
    Text = Text & "Input :" & vbLf
    Text = Text & "Single cell only." & vbLf
    Text = Text & "Cell reference supported." & vbLf
    Text = Text & "UDF result supported." & vbLf
    Text = Text & "Multiple cells are not supported." & vbLf
    Text = Text & "Processing :" & vbLf
    Text = Text & "Extracts numeric characters 0-9 and ０-９." & vbLf
    Text = Text & "Full-width numeric characters are converted to half-width." & vbLf
    Text = Text & "Kanji numerals are not extracted." & vbLf
    Text = Text & "Non-numeric characters are ignored." & vbLf
    Text = Text & "Output :" & vbLf
    Text = Text & "Numeric characters returned as text." & vbLf
    Text = Text & "#N/A if no numeric characters exist." & vbLf
    Text = Text & "Version : Ver2.0.0"
    HelpNumOnly = HelpArray(Text)
End Function

Private Function HelpNumSpill() As Variant
    Dim Text As String
    Text = "Function : NumSpill" & vbLf
    Text = Text & "Description : Extracts numeric sequences from text and returns them as text strings in a vertical spill." & vbLf
    Text = Text & "Syntax : =NumSpill(Target)" & vbLf
    Text = Text & "Input :" & vbLf
    Text = Text & "Single cell only." & vbLf
    Text = Text & "Cell reference supported." & vbLf
    Text = Text & "UDF result supported." & vbLf
    Text = Text & "Multiple cells are not supported." & vbLf
    Text = Text & "Processing :" & vbLf
    Text = Text & "Extracts consecutive numeric characters 0-9 and ０-９." & vbLf
    Text = Text & "Full-width numeric characters are converted to half-width." & vbLf
    Text = Text & "Kanji numerals are not extracted." & vbLf
    Text = Text & "Non-numeric characters separate numeric sequences." & vbLf
    Text = Text & "Output :" & vbLf
    Text = Text & "Numeric sequences returned as text strings." & vbLf
    Text = Text & "Vertical spill." & vbLf
    Text = Text & "#N/A if no numeric sequences exist." & vbLf
    Text = Text & "Version : Ver2.0.0"
    HelpNumSpill = HelpArray(Text)
End Function

Private Function HelpStrInfo() As Variant
    Dim Text As String
    Text = "Function : StrInfo" & vbLf
    Text = Text & "Syntax : =StrInfo(Target)" & vbLf
    Text = Text & "Description : Splits a string into Text and Number blocks and returns block information." & vbLf
    Text = Text & "Arguments :" & vbLf
    Text = Text & "Target: Target cell containing the string to analyze." & vbLf
    Text = Text & "Single cell only." & vbLf
    Text = Text & "Multiple cells return #VALUE!." & vbLf
    Text = Text & "Return :" & vbLf
    Text = Text & "Block, Start, Length, Type." & vbLf
    Text = Text & "Dynamic Array." & vbLf
    Text = Text & "Header included." & vbLf
    Text = Text & "Vertical spill." & vbLf
    Text = Text & "Block: Text or Number block." & vbLf
    Text = Text & "Start: Starting position in the original string." & vbLf
    Text = Text & "Length: Length in the original string." & vbLf
    Text = Text & "Type: Text or Number." & vbLf
    Text = Text & "Number :" & vbLf
    Text = Text & "Half-width and full-width digits 0-9 / ０-９." & vbLf
    Text = Text & "Full-width digits are converted to half-width digits in Block." & vbLf
    Text = Text & "Kanji numerals are treated as Text." & vbLf
    Text = Text & "Example : =StrInfo(A1)" & vbLf
    Text = Text & "Version : Ver2.0.0"
    HelpStrInfo = HelpArray(Text)
End Function

'==================================================
' Specification comments are unchanged
'==================================================


