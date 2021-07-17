# Export-ExcelFunction

## 概要

指定した Excel ファイルから Excel 関数を取り出します。

## 説明

Export-ExcelFunction 関数は指定した Excel ファイルから Excel 関数を抽出します。
主に、Excel 関数ごとの使用回数を調べるのに使うことになると思います。

Excel 関数を抽出する際、Export-ExcelFunction は対象のファイルをコピーし、ZIP ファイルにリネームし、XML ファイルを取得するために展開します。
そのため対象のファイルは .xlsx または .xlsm、.xlam、.xltx、.xltm 形式のファイルである必要があります。

Export-ExcelFunction は 'WorkbookIndex' を付けてExcel 関数を返します。同じ 'WorkbookIndex' が付いた関数は同じブックから抽出されたことを表します。

## 例 1

```ps1
Get-ChildItem -Filter *.xl?? -File | Export-ExcelFunction
```

上記のコマンドは入力されたファイルから Excel 関数を返します。結果は以下のようになります。

```
WorkbookIndex     Function
-------------     --------
20210625213459224 SUM
20210625213459224 IF
20210625213459224 AVERAGE
20210625213459224 SUM
20210625213401369 RAND
20210625213401369 RAND
20210625213401369 MAX
20210625213401369 MIN
20210625213402480 COUNTIF
20210625213402480 COUNTIF
20210625213402480 SUMIF
```

この場合、最初の4つの関数は WorkbookIndex が '20210625213459224' であるため、同じブックから抽出したことがわかります。

## 例 2

```ps1
$exportedFunctions = Get-ChildItem -Filter *.xl?? -File | Export-ExcelFunction

$measuredFunctions = $exportedFunctions | Group-Object -Property Function | 
    Select-Object `
        @{label="Function"; expression={$_.Name}}, 
        @{label="CountByCell"; expression={$_.Count}}, 
        @{label="CountByBook"; expression={@($_.Group | Select-Object -Property WorkbookIndex -Unique).Length}}

$measuredFunctions | Sort-Object -Property CountByBook -Descending | Select-Object -First 20
```

結果は以下のようになります。

```
Function CountByCell CountByBook
-------- ----------- -----------
IF              3037          15
IFERROR          360          10
SUM              110           8
ROW              193           8
INDEX            436           7
VLOOKUP          385           7
COUNTIF         2606           7
AND              153           6
RAND              21           5
MATCH            505           5
COLUMN           171           5
OR                59           4
WEEKDAY           12           3
OFFSET            26           3
RANK             111           3
CHOOSE            22           3
AVERAGE            4           3
RIGHT             35           2
INDIRECT           4           2
DATE               2           2
```