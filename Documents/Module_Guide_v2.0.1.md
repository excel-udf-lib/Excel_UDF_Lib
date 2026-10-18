# Excel UDF Lib Module Guide Ver2.0.x

## 1. Document Purpose

本書は、Excel UDF Lib Ver2.0.xで提供するVBA Module（.basファイル）の構成、導入方法、利用方法を説明する公開文書です。

Excel UDF Libでは、必要なModuleをExcel VBAプロジェクトへ追加することで、対応するUDFを利用できます。

モジュール構成については、本書の記載よりも、GitHubで現在公開されている`.bas`ファイルを優先します。

---

## 2. Module Structure

Excel UDF Lib Ver2.0.xでは、以下の5 Moduleを提供します。

| Module | Category | UDF数 |
|---|---|---:|
| modFile.bas | File系 | 8 |
| modSheet.bas | Sheet系 | 6 |
| modCell.bas | Cell系 | 9 |
| modTable.bas | Table系 | 2 |
| modColor.bas | Color系 | 5 |
| **合計** | | **30** |

### 2.1 modFile

File系UDFを提供します。

- FileInfo
- FileList
- FileExists
- FileName
- FilePath
- FileExt
- FileSize
- FileDate

### 2.2 modSheet

Sheet系UDFを提供します。

- SheetInfo
- SheetList
- SheetExists
- SheetName
- SheetIndex
- SheetCount

### 2.3 modCell

Cell系UDFを提供します。

- CellInfo
- CellType
- CellHasFormula
- CellFormula
- CellFormat
- CellBlank
- CellError
- CellLastRow
- CellLastColumn

### 2.4 modTable

Table系UDFを提供します。

- TableList
- TableInfo

### 2.5 modColor

Color系UDFを提供します。

- ColorInfo
- CellColorHex
- CellColorRGB
- FontColorHex
- FontColorRGB

---

## 3. Installation

Excel UDF LibのUDFを利用するには、必要な`.bas` ModuleをExcel VBAプロジェクトへ追加します。

### 3.1 VBA Editorを開く

Excelを起動し、`Alt + F11`を押してVBA Editorを開きます。

### 3.2 Moduleをインポートする

VBA Editorで以下を選択します。

`File` → `Import File`

利用する`.bas`ファイルを選択してください。

### 3.3 Moduleの追加を確認する

Project ExplorerにModuleが追加されます。

例：

VBAProject
└─ Modules
├─ modFile
├─ modSheet
├─ modCell
├─ modTable
└─ modColor

利用するUDFに対応するModuleを追加してください。

---

## 4. Usage

ModuleをExcel VBAプロジェクトへ追加すると、対応するUDFをExcelワークシート上で通常のExcel関数と同じように利用できます。

### 4.1 基本的な利用方法

Module追加後、ワークシート上でUDFを入力します。

例：

=FileExists("C:\Test\Test.xlsx")

=CellInfo(A1)

=TableInfo(A1)

### 4.2 Module単位での利用

Excel UDF Libでは、UDFをModule単位で提供しています。

利用したいUDFが含まれるModuleを追加してください。

ModuleとUDFの対応については、現在GitHubで公開されている`.bas`ファイルを基準とします。

---

## 5. Module Selection

Excel UDF Lib Ver2.0.xでは、以下の5 Moduleを提供しています。

### 5.1 modFile

**Category：File系**

提供UDF：

- FileInfo
- FileList
- FileExists
- FileName
- FilePath
- FileExt
- FileSize
- FileDate

### 5.2 modSheet

**Category：Sheet系**

提供UDF：

- SheetInfo
- SheetList
- SheetExists
- SheetName
- SheetIndex
- SheetCount

### 5.3 modCell

**Category：Cell系**

提供UDF：

- CellInfo
- CellType
- CellHasFormula
- CellFormula
- CellFormat
- CellBlank
- CellError
- CellLastRow
- CellLastColumn

### 5.4 modTable

**Category：Table系**

提供UDF：

- TableList
- TableInfo

### 5.5 modColor

**Category：Color系**

提供UDF：

- ColorInfo
- CellColorHex
- CellColorRGB
- FontColorHex
- FontColorRGB

---

## 6. UDF Classification

Excel UDF Libでは、Public UDFを以下の3種類に分類しています。

### 6.1 Info系UDF

複数の情報を一覧形式で取得するUDFです。

特徴：

- Header付き
- 表形式出力
- Dynamic Array対応
- Column方向Spill
- Address列を含む

対象：

- FileInfo
- SheetInfo
- CellInfo
- TableInfo
- ColorInfo

### 6.2 List系UDF

複数の対象を一覧表示するUDFです。

対象：

- FileList
- SheetList
- TableList

### 6.3 単機能系UDF

単一の情報取得や判定を行うUDFです。

対象：

- FileExists
- FileName
- FilePath
- FileExt
- FileSize
- FileDate
- SheetExists
- SheetName
- SheetIndex
- SheetCount
- CellType
- CellHasFormula
- CellFormula
- CellFormat
- CellBlank
- CellError
- CellLastRow
- CellLastColumn
- CellColorHex
- CellColorRGB
- FontColorHex
- FontColorRGB

---

## 7. Help Function

Excel UDF Libでは、Public UDFにHelp機能を提供しています。

Helpは、UDFの引数に`"?"`を指定して表示します。

例：

=FileInfo("?")

=CellInfo("?")

Helpでは、UDFの概要、引数、戻り値、使用方法などを確認できます。

Helpの詳細については、UDF Catalogを参照してください。

---

## 8. Excel Environment

Excel UDF Libは、VBAで作成されたUDFライブラリです。

利用には、VBAを使用できるExcel環境が必要です。

また、UDFによって必要なExcel機能が異なります。

特にInfo系UDFやList系UDFなど、Dynamic Array（スピル）を利用するUDFでは、対応するExcel環境が必要です。

利用するUDFの対応条件については、UDF Catalogを確認してください。

---

## 9. Error Handling

Excel UDF Libでは、UDFの処理中に取得できない情報や不正な対象が発生した場合に、エラー処理を行います。

基本的な考え方は以下のとおりです。

| 状態 | 基本的な結果 |
|---|---|
| 正常 | 取得した情報を返す |
| 取得不能 | エラーまたは該当する結果を返す |
| 不正な対象 | エラーを返す |
| Help指定 | Help情報を返す |

具体的なエラー処理は、UDFの仕様に従います。

---

## 10. Related Documents

Excel UDF Libの利用・仕様確認には、以下の公開文書を参照してください。

### README

Excel UDF Lib全体の概要、利用方法、特徴などを説明しています。

### Module Guide

VBA Moduleの構成、導入方法、利用方法を説明しています。

### UDF Catalog

公開UDFの一覧と各UDFの概要を説明しています。

### LICENSE

Excel UDF Libのライセンス条件を定義しています。

---

## 11. Version

本書の対象バージョン：

**Excel UDF Lib Ver2.0.x**

ModuleおよびUDFの構成については、本書の記載よりも、GitHubで現在公開されている`.bas`ファイルを優先します。

Excel UDF Lib Ver2.0.xでは、以下の5 Module、30 UDFを提供しています。

| Module | UDF数 |
|---|---:|
| modFile | 8 |
| modSheet | 6 |
| modCell | 9 |
| modTable | 2 |
| modColor | 5 |
| **Total** | **30** |

---

## 12. Disclaimer

Excel UDF Libは無保証で提供されています。

提供しているModuleおよびUDFについて、すべての環境での動作や特定目的への適合性を保証するものではありません。

Excel UDF Libを利用する場合は、利用者自身の判断と責任において使用してください。

Excel UDF Libの利用によって発生した問題や損害について、開発者はいかなる責任も負いません。

---

## Version

Excel UDF Lib Module Guide Ver2.0.x

Excel UDF Lib Ver2.0.xの公開Module構成を基準としています。

Module構成については、GitHubで現在公開されている`.bas`ファイルを優先します。

---

## End of Document