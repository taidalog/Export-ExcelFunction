# Export-ExcelFunction

## SYNOPSIS

Exports Excel Functions used in inputted Excel files.

## DESCRIPTION

Export-ExcelFunction function returns Excel functions used in inputted Excel files.
It will be used in order to, mainly, know the appearance of each Excel function.

When ectracting Excel functions, Export-ExcelFunction copies target files, renames them as .zip files, and expand them in oreder to uses XML files.
So the targert files need to be .xlsx, .xlsm or .xlam files.

Export-ExcelFunction returns Excel functions with 'WorkbookIndex', which tells that Functions with the same 'WorkbookIndex' were found in the same workbook.

## EXAMPLE 1

```ps1
Get-ChildItem -Filter *.xl?? -File | Export-ExcelFunction
```

The command above will return the Excel functions found in the input files, just as below:

```
WorkbookIndex     Function
-------------     --------
20210625213459224 SUM
20210625213459224 IF
20210625213459224 AVERAGE
20210625213459224 SUM
20210625213459224 SUM
```

## EXAMPLE 2

```ps1
$exportedFunctions = Get-ChildItem -Filter *.xl?? -File | Export-ExcelFunction

$measuredFunctions = $exportedFunctions | Group-Object -Property Function | 
    Select-Object `
        @{label="Function"; expression={$_.Name}}, 
        @{label="CountByCell"; expression={$_.Count}}, 
        @{label="CountByBook"; expression={@($_.Group | Select-Object -Property WorkbookIndex -Unique).Length}}

$measuredFunctions | Sort-Object -Property CountByBook -Descending | Select-Object -First 20
```

The result will be below:

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