Attribute VB_Name = "modTable"
Option Explicit

'==================================================
' Excel UDF Lib Ver2.0
' Module : modTable
' Table Information UDF Library
' Compact Final (修正版)
' 修正点: ThisWorkbook 参照を廃止し、呼び出し元ブック／引数で指定したブック・シートを優先して解決
'==================================================

Private Const HELP_MARK As String = "?"

'==================================================
' 共通ユーティリティ
'==================================================

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

Private Function IsHelp(ByVal Value As Variant) As Boolean
    IsHelp = (GetText(Value) = HELP_MARK)
End Function

' 呼び出し元の Workbook を返す（セルから呼ばれた場合はそのセルの属する Workbook）
Private Function GetCallerWorkbook() As Workbook
    On Error Resume Next
    If TypeName(Application.Caller) = "Range" Then
        Set GetCallerWorkbook = Application.Caller.Parent.Parent
    Else
        Set GetCallerWorkbook = Application.ActiveWorkbook
    End If
    On Error GoTo 0
End Function

' 指定文字列が開いているブック名（またはフルパス）に一致する場合、その Workbook を返す
Private Function FindWorkbookByName(ByVal nameOrFull As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase(wb.Name) = LCase(nameOrFull) Or LCase(wb.FullName) = LCase(nameOrFull) Then
            Set FindWorkbookByName = wb
            Exit Function
        End If
    Next wb
End Function

' 指定のワークブック内に指定シート名があれば True
Private Function SheetExistsInWorkbook(ByVal wb As Workbook, ByVal SheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SheetName)
    SheetExistsInWorkbook = Not ws Is Nothing
    On Error GoTo 0
End Function

'==================================================
' Help 出力ヘルパー
'==================================================

Private Function HelpArray( _
    ByVal FunctionName As String, _
    ByVal Description As String, _
    ByVal Syntax As String _
) As Variant

    Dim arr(1 To 4, 1 To 2) As Variant

    arr(1, 1) = "Function"
    arr(1, 2) = FunctionName

    arr(2, 1) = "Description"
    arr(2, 2) = Description

    arr(3, 1) = "Syntax"
    arr(3, 2) = Syntax

    arr(4, 1) = "Version"
    arr(4, 2) = "Excel UDF Lib Ver2.0"

    HelpArray = arr

End Function

'==================================================
' Public API
'==================================================

' TableList([Target])
' - Target が省略: 呼び出し元ブック内のすべてのテーブルを返す
' - Target がシート名（文字列）: 呼び出し元ブックのそのシート内テーブルを返す
' - Target がブック名（文字列）: その開いているブック内のすべてのテーブルを返す
' - Target が Range: その Range が属するシートのテーブルを返す
Public Function TableList( _
    Optional ByVal Target As Variant _
) As Variant

    Dim wb As Workbook
    Dim SheetName As String
    Dim wbFound As Workbook

    If IsHelp(Target) Then
        TableList = HelpArray( _
            "TableList", _
            "Returns table list.", _
            "=TableList([SheetOrWorkbook])")
        Exit Function
    End If

    If IsMissing(Target) Then
        Set wb = GetCallerWorkbook()
        If wb Is Nothing Then TableList = ReturnNA(): Exit Function
        TableList = BuildTableList(wb)
        Exit Function
    End If

    ' Range 指定
    If TypeName(Target) = "Range" Then
        Set wb = Target.Parent.Parent
        If wb Is Nothing Then TableList = ReturnNA(): Exit Function
        TableList = BuildTableList(wb)
        Exit Function
    End If

    ' 文字列指定: まず呼び出し元ブックのシート名として解釈を試みる
    SheetName = GetText(Target)
    Set wb = GetCallerWorkbook()
    If Not wb Is Nothing Then
        If SheetExistsInWorkbook(wb, SheetName) Then
            TableList = BuildSheetTableInfo(wb, SheetName)
            Exit Function
        End If
    End If

    ' 次に文字列をブック名として検索
    Set wbFound = FindWorkbookByName(SheetName)
    If Not wbFound Is Nothing Then
        TableList = BuildTableList(wbFound)
        Exit Function
    End If

    ' 解決できない場合
    TableList = ReturnNA()

End Function

' TableInfo([Target])
' - Target 省略: 呼び出し元ブック内のすべてのテーブル情報を返す
' - Target がシート名（文字列）: 呼び出し元ブックのそのシートのテーブルを返す
' - Target がブック名（文字列）: そのブック内のすべてのテーブルを返す
' - Target が Range: その Range が属するシートのテーブルを返す
Public Function TableInfo( _
    Optional ByVal Target As Variant _
) As Variant

    If IsHelp(Target) Then
        TableInfo = HelpArray( _
            "TableInfo", _
            "Returns table information.", _
            "=TableInfo([SheetOrWorkbook])")
        Exit Function
    End If

    If IsMissing(Target) Then
        Dim wb0 As Workbook
        Set wb0 = GetCallerWorkbook()
        If wb0 Is Nothing Then TableInfo = ReturnNA(): Exit Function
        TableInfo = BuildTableInfo(wb0)
        Exit Function
    End If

    If TypeName(Target) = "Range" Then
        Dim wbR As Workbook
        Dim shNameR As String
        Set wbR = Target.Parent.Parent
        shNameR = Target.Parent.Name
        If wbR Is Nothing Then TableInfo = ReturnNA(): Exit Function
        TableInfo = BuildSheetTableInfo(wbR, shNameR)
        Exit Function
    End If

    ' 文字列指定: まず呼び出し元ブックのシート名として解釈を試みる
    Dim txt As String
    txt = GetText(Target)

    Dim wbCaller As Workbook
    Set wbCaller = GetCallerWorkbook()
    If Not wbCaller Is Nothing Then
        If SheetExistsInWorkbook(wbCaller, txt) Then
            TableInfo = BuildSheetTableInfo(wbCaller, txt)
            Exit Function
        End If
    End If

    ' 次にブック名として検索
    Dim wbFound2 As Workbook
    Set wbFound2 = FindWorkbookByName(txt)
    If Not wbFound2 Is Nothing Then
        TableInfo = BuildTableInfo(wbFound2)
        Exit Function
    End If

    ' 見つからない場合
    TableInfo = ReturnNA()

End Function

'==================================================
' Core: Table ／ TableInfo ビルド（Workbook を引数に取る）
'==================================================

Private Function BuildTableList(ByVal wb As Workbook) As Variant
    Dim ws As Worksheet
    Dim table As ListObject
    Dim Count As Long
    Dim Row As Long
    Dim arr() As Variant

    If wb Is Nothing Then BuildTableList = ReturnNA(): Exit Function

    For Each ws In wb.Worksheets
        Count = Count + ws.ListObjects.Count
    Next ws

    If Count = 0 Then
        BuildTableList = ReturnNA()
        Exit Function
    End If

    ReDim arr(1 To Count + 1, 1 To 4)

    arr(1, 1) = "Name"
    arr(1, 2) = "Sheet"
    arr(1, 3) = "Address"
    arr(1, 4) = "ColumnsList"

    Row = 2

    For Each ws In wb.Worksheets
        For Each table In ws.ListObjects
            arr(Row, 1) = table.Name
            arr(Row, 2) = ws.Name
            arr(Row, 3) = table.Range.Address(False, False)
            arr(Row, 4) = GetColumnList(table)
            Row = Row + 1
        Next table
    Next ws

    BuildTableList = arr
End Function

Private Function BuildTableInfo(ByVal wb As Workbook) As Variant
    Dim ws As Worksheet
    Dim table As ListObject
    Dim Count As Long
    Dim Row As Long
    Dim arr() As Variant

    If wb Is Nothing Then BuildTableInfo = ReturnNA(): Exit Function

    For Each ws In wb.Worksheets
        Count = Count + ws.ListObjects.Count
    Next ws

    If Count = 0 Then
        BuildTableInfo = ReturnNA()
        Exit Function
    End If

    ReDim arr(1 To Count + 1, 1 To 7)

    SetTableInfoHeader arr

    Row = 2

    For Each ws In wb.Worksheets
        For Each table In ws.ListObjects
            SetTableInfoRow arr, Row, table
            Row = Row + 1
        Next table
    Next ws

    BuildTableInfo = arr
End Function

Private Function BuildSheetTableInfo( _
    ByVal wb As Workbook, _
    ByVal SheetName As String _
) As Variant

    Dim ws As Worksheet
    Dim table As ListObject
    Dim Count As Long
    Dim Row As Long
    Dim arr() As Variant

    On Error Resume Next
    Set ws = wb.Worksheets(SheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        BuildSheetTableInfo = ReturnNA()
        Exit Function
    End If

    Count = ws.ListObjects.Count
    If Count = 0 Then
        BuildSheetTableInfo = ReturnNA()
        Exit Function
    End If

    ReDim arr(1 To Count + 1, 1 To 7)

    SetTableInfoHeader arr

    Row = 2

    For Each table In ws.ListObjects
        SetTableInfoRow arr, Row, table
        Row = Row + 1
    Next table

    BuildSheetTableInfo = arr
End Function

'==================================================
' Target 解決ユーティリティ（文字列 / Range をサポート）
'==================================================

Private Function GetTargetSheetName( _
    ByVal Target As Variant _
) As String

    If IsObject(Target) Then
        If TypeName(Target) = "Range" Then
            If Target.Cells.Count = 1 Then
                GetTargetSheetName = GetText(Target.Value)
            End If
        End If
        Exit Function
    End If

    GetTargetSheetName = GetText(Target)
End Function

'==================================================
' TableInfo の出力整形ヘルパー
'==================================================

Private Sub SetTableInfoHeader( _
    ByRef arr As Variant _
)
    arr(1, 1) = "Name"
    arr(1, 2) = "Sheet"
    arr(1, 3) = "Address"
    arr(1, 4) = "Rows"
    arr(1, 5) = "Columns"
    arr(1, 6) = "Index"
    arr(1, 7) = "ColumnsList"
End Sub

Private Sub SetTableInfoRow( _
    ByRef arr As Variant, _
    ByVal Row As Long, _
    ByVal table As ListObject _
)
    arr(Row, 1) = table.Name
    arr(Row, 2) = table.Parent.Name
    arr(Row, 3) = table.Range.Address(False, False)
    arr(Row, 4) = GetTableRows(table)
    arr(Row, 5) = table.ListColumns.Count
    arr(Row, 6) = GetTableIndex(table)
    arr(Row, 7) = GetColumnList(table)
End Sub

'==================================================
' Getter ヘルパー
'==================================================

Private Function GetTableRows( _
    ByVal table As ListObject _
) As Long
    If Not table.DataBodyRange Is Nothing Then
        GetTableRows = table.DataBodyRange.Rows.Count
    End If
End Function

Private Function GetTableIndex( _
    ByVal table As ListObject _
) As Long
    Dim t As ListObject
    Dim index As Long
    index = 1
    For Each t In table.Parent.ListObjects
        If t.Name = table.Name Then
            GetTableIndex = index
            Exit Function
        End If
        index = index + 1
    Next t
End Function

Private Function GetColumnList( _
    ByVal table As ListObject _
) As String
    Dim col As ListColumn
    Dim Result As String
    For Each col In table.ListColumns
        If Result <> "" Then
            Result = Result & ","
        End If
        Result = Result & col.Name
    Next col
    GetColumnList = Result
End Function


