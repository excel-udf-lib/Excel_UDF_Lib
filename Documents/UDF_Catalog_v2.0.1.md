# Excel UDF Lib UDF Catalog Ver2.0.x

## 1. Document Purpose

本書は、Excel UDF Lib Ver2.0.xで公開しているUDFの一覧および概要を説明する公開文書です。

Excel UDF Libでは、Excel標準関数では取得しにくい情報や、Excel作業で発生する「少し面倒」を解決するためのUDFを提供しています。

本書は、公開UDFの名称、分類、用途を確認するためのカタログです。

UDFおよびModuleの構成については、本書の記載よりも、GitHubで現在公開されている`.bas`ファイルを優先します。

---

## 2. UDF Classification

Excel UDF Libでは、Public UDFを以下の3種類に分類しています。

### 2.1 Info系

複数の情報を一覧形式で取得するUDFです。

- FileInfo
- SheetInfo
- CellInfo
- TableInfo
- ColorInfo

### 2.2 List系

対象となる項目を一覧形式で取得するUDFです。

- FileList
- SheetList
- TableList

### 2.3 単機能系

単一の情報取得や判定を行うUDFです。

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

## 3. File系UDF

ファイルに関する情報の取得や確認を行うUDFです。

### 3.1 FileInfo

ファイルに関する複数の情報を一覧形式で取得します。

### 3.2 FileList

ファイルを一覧形式で取得します。

### 3.3 FileExists

指定したファイルが存在するかを確認します。

### 3.4 FileName

指定したファイルのファイル名を取得します。

### 3.5 FilePath

指定したファイルのパスを取得します。

### 3.6 FileExt

指定したファイルの拡張子を取得します。

### 3.7 FileSize

指定したファイルのサイズを取得します。

### 3.8 FileDate

指定したファイルの日付情報を取得します。

---

## 4. Sheet系UDF

Excelワークブック内のワークシートに関する情報の取得や確認を行うUDFです。

### 4.1 SheetInfo

ワークシートに関する複数の情報を一覧形式で取得します。

### 4.2 SheetList

ワークブック内のワークシートを一覧形式で取得します。

### 4.3 SheetExists

指定したワークシートが存在するかを確認します。

### 4.4 SheetName

ワークシートの名前を取得します。

### 4.5 SheetIndex

ワークシートの番号を取得します。

### 4.6 SheetCount

ワークブック内のワークシート数を取得します。

---

## 5. Cell系UDF

セルの状態や情報を取得・確認するUDFです。

### 5.1 CellInfo

セルに関する複数の情報を一覧形式で取得します。

### 5.2 CellType

セルのデータ型を取得します。

### 5.3 CellHasFormula

セルに数式が設定されているかを確認します。

### 5.4 CellFormula

セルに設定されている数式を取得します。

### 5.5 CellFormat

セルの表示形式を取得します。

### 5.6 CellBlank

セルが空白かどうかを確認します。

### 5.7 CellError

セルがエラー状態かどうかを確認します。

### 5.8 CellLastRow

対象範囲の最終行を取得します。

### 5.9 CellLastColumn

対象範囲の最終列を取得します。

---

## 6. Table系UDF

Excelテーブルに関する情報を取得するUDFです。

### 6.1 TableList

Excelテーブルを一覧形式で取得します。

### 6.2 TableInfo

Excelテーブルに関する複数の情報を一覧形式で取得します。

---

## 7. Color系UDF

セルの背景色や文字色に関する情報を取得するUDFです。

### 7.1 ColorInfo

セルの色情報を一覧形式で取得します。

### 7.2 CellColorHex

セルの背景色をHex形式で取得します。

### 7.3 CellColorRGB

セルの背景色をRGB形式で取得します。

### 7.4 FontColorHex

セルの文字色をHex形式で取得します。

### 7.5 FontColorRGB

セルの文字色をRGB形式で取得します。

---

## 8. UDF Summary

Excel UDF Lib Ver2.0.0では、5 Module、30 UDFを提供しています。

| Module | Category | Info | List | 単機能 | UDF数 |
|---|---|---:|---:|---:|---:|
| modFile | File系 | 1 | 1 | 6 | 8 |
| modSheet | Sheet系 | 1 | 1 | 4 | 6 |
| modCell | Cell系 | 1 | 0 | 8 | 9 |
| modTable | Table系 | 1 | 1 | 0 | 2 |
| modColor | Color系 | 1 | 0 | 4 | 5 |
| **合計** | | **5** | **3** | **22** | **30** |

---

## 9. Help

Excel UDF Libでは、Public UDFにHelp機能を提供しています。

UDFの引数に`"?"`を指定することで、Help情報を確認できます。

例：

=FileInfo("?")

=CellInfo("?")

Helpでは、UDFの概要、引数、戻り値、使用方法などを確認できます。

---

## 10. Module

Excel UDF Lib Ver2.0.xでは、以下の5 Moduleを提供しています。

| Module | Category | UDF数 |
|---|---|---:|
| modFile | File系 | 8 |
| modSheet | Sheet系 | 6 |
| modCell | Cell系 | 9 |
| modTable | Table系 | 2 |
| modColor | Color系 | 5 |
| **Total** | | **30** |

各Moduleの導入方法については、Module Guideを参照してください。

---

## 11. Related Documents

Excel UDF Libの利用に必要な公開文書は以下のとおりです。

### README

Excel UDF Lib全体の概要、特徴、利用方法などを説明しています。

### Module Guide

VBA Moduleの構成、導入方法、利用方法を説明しています。

### UDF Catalog

公開しているUDFの一覧と概要を説明しています。

### LICENSE

Excel UDF Libのライセンス条件を定義しています。

---

## 12. Version

本書の対象バージョン：

Excel UDF Lib Ver2.0.x

本書は、Excel UDF Lib Ver2.0.xの公開UDF構成を基準としています。

UDFおよびModuleの構成については、本書の記載よりも、GitHubで現在公開されている`.bas`ファイルを優先します。

---

## End of Document

Excel UDF Lib UDF Catalog Ver2.0.x
