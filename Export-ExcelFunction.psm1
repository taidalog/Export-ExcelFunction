Set-StrictMode -Version Latest

function Export-ExcelFunction {
    <#
    .SYNOPSIS
    Exports Excel Functions used in inputted Excel files.

    .DESCRIPTION
    Export-ExcelFunction function returns Excel functions used in inputted Excel files.
    It will be used in order to, mainly, know the appearance of each Excel function.
    
    When ectracting Excel functions, Export-ExcelFunction copies target files, renames them as .zip files, and expand them in oreder to uses XML files.
    So the targert files need to be .xlsx, .xlsm or .xlam files.

    Export-ExcelFunction returns Excel functions with 'WorkbookIndex', which tells that Functions with the same 'WorkbookIndex' were found in the same workbook.

    .EXAMPLE
    Get-ChildItem -Filter *.xl?? -File | Export-ExcelFunction

    The command above will return the Excel functions found in the input files, just as below:

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

    In this case, first four Functions, whose WorkbookIndex are '20210625213459224', were from the same workbook.

    .EXAMPLE
    $exportedFunctions = Get-ChildItem -Filter *.xl?? -File | Export-ExcelFunction
    
    $measuredFunctions = $exportedFunctions | Group-Object -Property Function | 
        Select-Object `
            @{label="Function"; expression={$_.Name}}, 
            @{label="CountByCell"; expression={$_.Count}}, 
            @{label="CountByBook"; expression={@($_.Group | Select-Object -Property WorkbookIndex -Unique).Length}}
    
    $measuredFunctions | Sort-Object -Property CountByBook -Descending | Select-Object -First 20
    
    The result will be below:

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
    
    .PARAMETER Path
    .PARAMETER KeepWorkFile
    .INPUTS
    .OUTPUTS
    .LINK
    #>

    [CmdletBinding()]
    param (
        # Specifies a path to one or more locations.
        [Parameter(Mandatory=$true,
                   Position=0,
                   ParameterSetName="Path",
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Path to one or more locations.")]
        [Alias("PSPath")]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({Test-Path $_ -PathType Leaf})]
        [string[]]
        $Path,

        [switch]
        $KeepWorkFile
    )
    
    begin {
        # working directory
        $workingPath = Join-Path $PSScriptRoot 'work'

        # zip extension
        $zipExtension = '.zip'

        $xmlPathStructure = 'xl\worksheets'
        
        $regexPatternForFormula = '(?<=<f[^<]*>).+?(?=</f.*>)'
        $regexObjectForFormula = [regex]::new($regexPatternForFormula)
        
        $regexPatternForFunction = '(?<=[=\+\-\*/,&\(]*)[0-9A-Za-z\.]+(?=\()'
        $regexObjectForFunction = [regex]::new($regexPatternForFunction)

#        $workbookIndex = 0
    }
    
    process {
        # accepts Get-ChildItem -Filter *.xl?? -Recurse -File
        foreach ($p in $Path) {
            $convertedPath = Convert-Path $p
            Write-Verbose "convertedPath = $($convertedPath)"
            
#            $workbookIndex++
            [string]$workbookIndex = (Get-Date).ToString('yyyyMMddHHmmssfff')

            # path to the place to copy xl?? file
            $originalExtension = [System.IO.Path]::GetExtension($convertedPath)
            $newPath = $convertedPath.Replace($originalExtension, $zipExtension)
            $newFullName = Join-Path $workingPath (Split-Path $newPath -NoQualifier)
            Write-Verbose "newFullName   = $($newFullName)"
            
            if ((Test-Path (Split-Path $newFullName -Parent)) -eq $false) {
                New-Item (Split-Path $newFullName -Parent) -ItemType Directory | Out-Null
            }

            Copy-Item -LiteralPath $convertedPath -Destination $newFullName
            
            $expandDestinationPath = $newFullName.Replace($zipExtension, '')
            Write-Verbose "expandDestinationPath = $($expandDestinationPath)"

            try {
                Expand-Archive -LiteralPath $newFullName -DestinationPath $expandDestinationPath
                
                $xmlFilesPath = Join-Path $expandDestinationPath $xmlPathStructure
                Write-Verbose "xmlFilesPath          = $($xmlFilesPath)"

                $xmlFiles = Get-ChildItem -LiteralPath $xmlFilesPath -Filter *.xml -File
                
                # get formulas and functions
                foreach ($xf in $xmlFiles) {
                    $xmlContent = Get-Content -LiteralPath $xf.FullName
                    $matchedFormulas = $regexObjectForFormula.Matches($xmlContent)

                    foreach ($fml in $matchedFormulas) {
                        $matchedFunctions = $regexObjectForFunction.Matches($fml.Value)

                        foreach ($fnc in $matchedfunctions) {
                            [PSCustomObject]@{
                                WorkbookIndex = $workbookIndex;
                                Function = $fnc.Value
                            }
                        }
                    }
                    
                }
            }
            catch [System.Management.Automation.MethodInvocationException] {
                Write-Verbose "Workbook with password: $($expandDestinationPath)"
            }
            finally {
                if ($KeepWorkFile -eq $false) {
                    Remove-Item -LiteralPath $newFullName -Force -Recurse
                    Remove-Item -LiteralPath $expandDestinationPath -Force -Recurse
                }
            }

        }
    }
    
    end {
        if ($KeepWorkFile -eq $false) {
            Remove-Item -LiteralPath $workingPath -Recurse
        }
    }
}

Set-Alias -Name exexf -Value Export-ExcelFunction

Export-ModuleMember -Function * -Alias *