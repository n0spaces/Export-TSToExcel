<# TSToExcel.ps1

Copyright (c) 2021 Matt Schwartz

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

#>

<#
.SYNOPSIS
Export a Configuration Manager task sequence to an Excel sheet for documentation.

.DESCRIPTION
This function generates an Excel sheet from a given Configuration Manager task sequence. It provides customizable and
easy-to-navigate documentation, especially for large task sequences.

.PARAMETER Xml
The task sequence XML data.

.PARAMETER XmlPath
Path to an exported task sequence XML file.

.PARAMETER TaskSequence
A task sequence object obtained from the Get-CMTaskSequence cmdlet. Accepts pipeline.

.PARAMETER ExportPath
The path that the exported Excel file will be saved to. Must end in .xlsx, or .xlsm for macro-enabled files. If the
path isn't specified, Excel will be visible after the sheet is generated so you can save it manually.

.PARAMETER TSName
The name of the task sequence. This will appear at the top of the Excel sheet. If -TaskSequence is specified, then the
name is obtained from that instead.

.PARAMETER Show
Shows the Excel window after the sheet is generated.

.PARAMETER Macro
Adds macro-enabled buttons to the sheet that expand/collapse grouped task sequence steps. See notes for more info.

.PARAMETER Outline
Groups (outlines) rows that belong to the same task sequence group, so they can be expanded/collapsed. This is uglier
than the macro buttons (especially for nested groups), but doesn't required macro permissions.

.PARAMETER HideProgress
Hides the progress bar in the PowerShell window.

.EXAMPLE
PS>Get-CMTaskSequence -name "Windows 10 Image" | Export-TSToExcel -Show

Generate an Excel sheet from the task sequence named "Windows 10 Image" and show it.

.EXAMPLE
PS>$ts = Get-CMTaskSequence -Name "Task Sequence"
PS>Export-TSToExcel -TaskSequence $ts -ExportPath C:\ts.xlsx

Generate an Excel sheet from the task sequence and save it to C:\ts.xlsx.

.EXAMPLE
PS>$sequence = (Get-CMTaskSequence -Name "Task Sequence").Sequence  # XML string
PS>Export-TSToExcel -Xml $sequence -TSName "Task Sequence" -Macro -ExportPath C:\ts.xlsm -Show

Generate an Excel sheet from the given task sequence XML data with macro buttons, save it to C:\ts.xlsm, and show it.

.EXAMPLE
PS>(Get-CMTaskSequence -Name "Task Sequence").Sequence | Out-File C:\ts.xml
PS>Export-TSToExcel -XmlPath C:\ts.xml -TSName "Task Sequence" -Outline -ExportPath C:\ts.xlsx -HideProgress

Generate an Excel sheet from an exported task sequence XML file with outlined groups, save it to C:\ts.xlsx, and hide
the progress bar.

.NOTES
If macro buttons are used, this script will require access to the VBA object model. If access is not allowed (which is
the default behavior), the script will prompt if it's okay to change a registry setting to allow access. You can also
set this manually in Excel by navigating to File > Options > Trust Center > Trust Center Settings... > Macro Settings,
and checking "Trust access to the VBA project object model". **This will allow scripts and programs to modify and run
macro scripts in Excel. Use wisely.**

.LINK
https://github.com/n0spaces/Export-TSToExcel
#>
function Export-TSToExcel
{
    param (
        [Parameter(ParameterSetName="FromXml", Mandatory)]
        [ValidateNotNullOrEmpty()]
        [xml] $Xml,

        [Parameter(ParameterSetName="FromXmlPath", Mandatory)]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo] $XmlPath,

        [Parameter(ParameterSetName="FromTaskSequence", Mandatory, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [object] $TaskSequence,

        [Parameter(ParameterSetName="FromTaskSequence")]
        [Parameter(ParameterSetName="FromXml")]
        [Parameter(ParameterSetName="FromXmlPath")]
        [System.IO.FileInfo] $ExportPath,

        [Parameter(ParameterSetName="FromXml")]
        [Parameter(ParameterSetName="FromXmlPath")]
        [string] $TSName = "Task Sequence",

        [Parameter(ParameterSetName="FromTaskSequence")]
        [Parameter(ParameterSetName="FromXml")]
        [Parameter(ParameterSetName="FromXmlPath")]
        [switch] $Show,

        [Parameter(ParameterSetName="FromTaskSequence")]
        [Parameter(ParameterSetName="FromXml")]
        [Parameter(ParameterSetName="FromXmlPath")]
        [switch] $Macro,

        [Parameter(ParameterSetName="FromTaskSequence")]
        [Parameter(ParameterSetName="FromXml")]
        [Parameter(ParameterSetName="FromXmlPath")]
        [switch] $Outline,
        
        [Parameter(ParameterSetName="FromTaskSequence")]
        [Parameter(ParameterSetName="FromXml")]
        [Parameter(ParameterSetName="FromXmlPath")]
        [switch] $HideProgress
    )

    try
    {
        Set-Variable -Name TSName -Option AllScope
        Set-Variable -Name LastUpdated -Option AllScope
        Set-Variable -Name Sequence -Option AllScope

        if ($null -ne $TaskSequence) {
            $TSName = $TaskSequence.Name
            $LastUpdated = $TaskSequence.LastRefreshTime
            $Sequence = ([xml]($TaskSequence.Sequence)).sequence
        }
        elseif ($null -ne $Xml) {
            $LastUpdated = Get-Date
            $Sequence = $Xml
            if ($null -ne $Xml.sequence) {
                $Sequence = $Xml.sequence
            }
        }
        elseif ($null -ne $XmlPath) {
            if (-not $XmlPath.Exists) {
                throw "$XmlPath does not exist."
            }
            $LastUpdated = Get-Date
            $Sequence = ([Xml](Get-Content $XmlPath)).sequence
        }

        # error checking
        if ($null -ne $ExportPath) {
            # not an excel file
            if ($ExportPath.Extension -ne ".xlsx" -and $ExportPath.Extension -ne ".xlsm") {
                throw "The specified export path does not appear to be an Excel file. Please make sure the path ends with the .xlsx or .xlsm extensions."
            }
            # macro is used, but path is not macro-enabled
            if ($Macro -and $ExportPath.Extension -eq ".xlsx") {
                throw "The -Macro switch was used, but the path is not for a macro-enabled file. Please use a file path that ends with .xlsm, or omit the -Macro switch."
            }
            # macro is not used, but path is macro-enabled
            if ($ExportPath.Extension -eq ".xlsm" -and -not $Macro) {
                Write-Warning "The path is for a macro-enabled file, but the -Macro tag was not used. The file will not have macros."
            }
        }
        # no export path
        elseif (-not $Show) {
            Write-Warning "No path was specified. Excel will be visible after the sheet is generated so you can save it."
            $Show = $true
        }

        $ColorGroup = 0xE7C6B4
        $ColorStep = 0x99E6FF
        $ColorGroupDisabled = 0xC9C9C9
        $ColorStepDisabled = 0xDBDBDB

        # VBA docs apply to this com object: https://docs.microsoft.com/en-us/office/vba/api/overview/excel
        $excel = New-Object -ComObject Excel.Application

        # vba security settings
        if ($Macro -and (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\Excel\Security").AccessVBOM -ne 1) {
            $q = "To use macro buttons, the following registry change will occur:`n" +
            "    HKCU\Software\Microsoft\Office\$($excel.Version)\Excel\Security\AccessVBOM = 1`n" +
            "This will give scripts access to the VBA object model, so they can add and execute VBA code in documents. Is this okay?`n" +
            "You can also set this in Excel by going to File > Options > Trust Center > Trust Center Settings > Macro Settings.`n" +
            "If you prefer to not allow VBA access, you can use -Outline instead of -Macro to group rows together."

            $choices = "&Yes", "&No"
            $result = $Host.UI.PromptForChoice("Access VBA Object Model", $q, $choices, 1)
            switch ($result) {
                0 {
                    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\Excel\Security" -Name AccessVBOM -Value 1 -Force | Out-Null
                }
                Default { return }
            }
        }
        Write-Host $ResetRegistryVBOM

        $wb = $excel.Workbooks.Add()
        $ws = $wb.ActiveSheet

        $excel.DisplayAlerts = $false

        if (-not $HideProgress) {
            Write-Progress -Activity "Generating Excel sheet..." -Status "Preparing..." -PercentComplete 0
        }

        # first row
        $text = "$TSName (Last updated: $($LastUpdated.ToShortDateString()) $($LastUpdated.ToShortTimeString()))"
        $ws.Range("A1").Cells = $text
        $ws.Range("A1").Cells.Font.Bold = $true
        $ws.Range("A1").Cells.Font.Size = 14
        $ws.Range("A1:F1").Merge()

        # headers
        $ws.Range("A2").Cells = "Name"
        $ws.Range("B2").Cells = "Type"
        $ws.Range("C2").Cells = "Description"
        $ws.Range("D2").Cells = "Conditions"
        $ws.Range("E2").Cells = "Continue on Error"
        $ws.Range("F2").Cells = "Settings"
        $ws.Rows("2").Font.Bold = $true

        # sheet formatting
        $ws.Cells.VerticalAlignment = -4160 # top alignment
        $ws.Columns("C:E").WrapText = $true

        Set-Variable -Name CurrentRow -Option AllScope -Value 2

        $outer = $Sequence.OuterXml
        $TotalEntries = ([regex]"<step").Matches($outer).Count + ([regex]"<group").Matches($outer).Count

        if ($Macro) {
            $Module = $wb.VBProject.VBComponents.Add(1)
            $VbaModule =
"Sub ToggleRowsHidden(rowsRange As String, triangle As Shape)
Dim Rows As Range
Set Rows = ActiveSheet.Rows(rowsRange)
If Rows.Hidden Then
    Rows.Hidden = False
    triangle.Rotation = 180
Else
    Rows.Hidden = True
    triangle.Rotation = 90
End If
End Sub"
            $Module.CodeModule.AddFromString($VbaModule)
        }

        function ConvertOperator {
            param ([string]$Operator)
            switch ($operator) {
                "equals" { return "=" }
                "notEquals" { return "!=" }
                "notExists" { return "does not exist"}
                "greater" { return ">" }
                "greaterEqual" { return ">=" }
                "less" { return "<" }
                "lessEqual" { return "<=" }
                default { return "$operator" }
            }
        }

        function ConvertTimeStamp {
            param ([string]$TimeStamp)
            $dt = [datetime]::ParseExact($TimeStamp.Substring(0, 18), "yyyyMMddHHmmss.fff", $null)
            return "$($dt.ToShortDateString()) $($dt.ToLongTimeString())"
        }
        function ParseCondition
        {
            param (
                $Element,
                $IndentLevel = 0
            )
            
            $text = ""
            $Indent = "    " * $IndentLevel

            # no conditions
            if ($null -eq $Element) {
                return $text
            }

            # root condition element
            if ($Element.LocalName -eq "condition") {
                foreach ($Child in $Element.ChildNodes) {
                    $text += ParseCondition -Element $Child
                }
                $text = $text.TrimEnd()
            }
            
            # operator element with child expressions
            elseif ($Element.LocalName -eq "operator") {
                $text += $Indent
                switch ($Element.type) {
                    "and" { $text += "All are true:`n" }
                    "or" { $text += "Any are true:`n" }
                    "not" { $text += "None are true:`n" }
                    Default { $text += "$($Element.type)`n"}
                }
                foreach ($Child in $Element.ChildNodes) {
                    $text += ParseCondition -Element $Child -IndentLevel ($IndentLevel + 1)
                }
            }

            # expression element
            elseif ($Element.LocalName -eq "expression") {
                $expression = "$Indent"
                switch ($Element.type) {
                    "SMS_TaskSequence_VariableConditionExpression" {
                        $variable = ($Element.variable | Where-Object name -eq "Variable").InnerText
                        $operator = ($Element.variable | Where-Object name -eq "Operator").InnerText
                        $value = ($Element.variable | Where-Object name -eq "Value").InnerText

                        $expression += "Variable $variable $(ConvertOperator $operator) "
                        if ($value -ne "" -and $null -ne $value) {
                            $expression += " `"$value`""
                        }
                    }
                    "SMS_TaskSequence_FolderConditionExpression" {
                        $path = ($Element.variable | Where-Object name -eq "Path").InnerText
                        $dt = ($Element.variable | Where-Object name -eq "DateTime").InnerText
                        $operator = ($Element.variable | Where-Object name -eq "DateTimeOperator").InnerText
                        $expression += "Folder `"$path`" exists"
                        if ($null -ne $dt) {
                            $expression += " and timestamp $(ConvertOperator $operator) $(ConvertTimeStamp $dt)"
                        }
                    }
                    "SMS_TaskSequence_FileConditionExpression" {
                        $path = ($Element.variable | Where-Object name -eq "Path").InnerText
                        $dt = ($Element.variable | Where-Object name -eq "DateTime").InnerText
                        $dtOperator = ($Element.variable | Where-Object name -eq "DateTimeOperator").InnerText
                        $version = ($Element.variable | Where-Object name -eq "Version").InnerText
                        $verOperator = ($Element.variable | Where-Object name -eq "VersionOperator").InnerText
                        $expression += "File `"$path`" exists"
                        if ($null -ne $dt) {
                            $expression += ", timestamp $(ConvertOperator $dtOperator) $(ConvertTimeStamp $dt)"
                        }
                        if ($null -ne $version) {
                            $expression += ", version $(ConvertOperator $verOperator) $version"
                        }
                    }
                    "SMS_TaskSequence_WMIConditionExpression" {
                        $ns = ($Element.variable | Where-Object name -eq "Namespace").InnerText
                        $q = ($Element.variable | Where-Object name -eq "Query").InnerText
                        $expression += "WMI Namespace: `"$ns`" Query: `"$q`""
                    }
                    "SMS_TaskSequence_RegistryConditionExpression" {
                        $data = ($Element.variable | Where-Object name -eq "Data").InnerText
                        $keypath = ($Element.variable | Where-Object name -eq "KeyPath").InnerText
                        $operator = ($Element.variable | Where-Object name -eq "Operator").InnerText
                        $type = ($Element.variable | Where-Object name -eq "Type").InnerText
                        $value = ($Element.variable | Where-Object name -eq "Value").InnerText
                        $expression += "Registry `"$keypath\$value`" ($type) $(ConvertOperator $operator) `"$data`""
                    }
                    Default { $expression += $Element.type }
                }
                $text += "$expression`n"
            }

            return $text
        }

        # convert a PascalCase sequence type to a space-delimited string
        function GetSequenceTypeFriendlyName
        {
            param ($Type)
            $Type = $Type.Replace("SMS_TaskSequence_", "").Replace("Action", "")
            switch ($Type) {
                # some types need to be hard-coded (otherwise PowerShell and BitLocker become Power Shell and Bit Locker)
                "RunPowerShellScript" { return "Run PowerShell Script" }
                "DisableBitLokcer" { return "Disable BitLocker" }
                "EnableBitLocker" { return "Enable BitLocker" }
                "OfflineEnableBitLocker" { return "Pre-provision BitLocker" }
                "AutoApply" { return "Auto Apply Drivers" }
                Default { return [regex]::Replace($Type, "([a-z](?=[A-Z])|[A-Z](?=[A-Z][a-z]))", "`$1 ") }
            }
        }

        # write a task sequence group or step to the excel sheet
        # called recursively for steps inside a group
        function WriteEntry
        {
            param (
                $Entry,
                $IndentLevel = 0,
                $Disabled = $false
            )

            if (-not $HideProgress) {
                [int]$p = (($CurrentRow - 1)/$TotalEntries) * 100
                Write-Progress -Activity "Generating Excel sheet..." -Status "Entry $($CurrentRow - 1)/$TotalEntries ($p%)" -PercentComplete $p
            }

            # all
            $CurrentRow++
            $ws.Range("A$CurrentRow").Cells = $Entry.GetAttribute("name")
            $ws.Range("C$CurrentRow").Cells = $Entry.GetAttribute("description")
            $ws.Range("D$CurrentRow").Cells = ParseCondition -Element $Entry.condition

            if ($Entry.GetAttribute("disable") -eq "true") {
                $Disabled = $true
            }

            if ($Entry.GetAttribute("continueOnError") -eq "true") {
                $ws.Range("E$CurrentRow").Cells = "Yes"
            }

            # steps
            if ($Entry.LocalName -eq "step")
            {
                $ws.Range("A$CurrentRow").IndentLevel = $IndentLevel
                if ($Disabled) {
                    $ws.Range("A$($CurrentRow):F$CurrentRow").Interior.Color = $ColorStepDisabled
                    $ws.Range("A$($CurrentRow):F$CurrentRow").Font.Strikethrough = $true
                } else {
                    $ws.Range("A$($CurrentRow):F$CurrentRow").Interior.Color = $ColorStep
                }

                $FriendlyType = GetSequenceTypeFriendlyName $Entry.GetAttribute("type")
                $ws.Range("B$CurrentRow").Cells = $FriendlyType

                # build variable list
                $vartext = ""
                foreach ($Variable in $Entry.defaultVarList.variable)
                {
                    $vartext += "$( $Variable.property ) = $( $Variable.InnerText )`n"
                }
                $ws.Range("F$CurrentRow").Cells = $vartext.TrimEnd()
            }

            # groups
            elseif ($Entry.LocalName -eq "group")
            {
                if ($Macro) {
                    $ws.Range("A$CurrentRow").IndentLevel = $IndentLevel + 1
                } else {
                    $ws.Range("A$CurrentRow").IndentLevel = $IndentLevel
                }

                if ($Disabled) {
                    $ws.Range("A$($CurrentRow):F$CurrentRow").Interior.Color = $ColorGroupDisabled
                    $ws.Range("A$($CurrentRow):F$CurrentRow").Font.Strikethrough = $true
                } else {
                    $ws.Range("A$($CurrentRow):F$CurrentRow").Interior.Color = $ColorGroup
                }
                $ws.Range("B$CurrentRow").Cells = "Group"
                $ws.Range("A$( $CurrentRow ):B$( $CurrentRow )").Font.Bold = $true

                # add expand button
                if ($Macro) {
                    $top = $ws.Range("A$CurrentRow").Top + 4
                    $left = (($IndentLevel - 1) * 7) + 4
                    $shape = $ws.Shapes.AddShape(7, $left, $top, 7, 7)
                    $shape.Fill.ForeColor.RGB = 0
                    $shape.Line.ForeColor.RGB = 0
                    $shape.Rotation = 180
                    $shape.Name = "ExpandShape$CurrentRow"
                }

                # recursively call for each child of this group
                $FirstRow = $CurrentRow + 1
                foreach ($Child in $Entry.ChildNodes)
                {
                    if ($Child.LocalName -eq "group" -or $Child.LocalName -eq "step")
                    {
                        WriteEntry -Entry $Child -IndentLevel ($IndentLevel + 1) -Disabled $Disabled
                    }
                }

                # outline (group) rows
                if ($Outline) {
                    $ws.Rows("$($FirstRow):$CurrentRow").Group() | Out-Null
                }

                # add code for button
                if ($Macro) {
                    $SubName = "$($shape.Name)Clicked"
                    $Code = "Sub $($SubName)()`n"
                    $Code += "ToggleRowsHidden `"$($FirstRow):$CurrentRow`", ActiveSheet.Shapes(`"$($shape.Name)`")`n"
                    $Code += "End Sub"
                    $Module.CodeModule.AddFromString($Code)
                    $shape.OnAction = $SubName
                }
            }
        }

        # fill entries
        $IndentLevel = 0
        if ($Macro) {
            $IndentLevel = 1
        }
        foreach ($Child in $Sequence.ChildNodes)
        {
            if ($Child.LocalName -eq "group" -or $Child.LocalName -eq "step")
            {
                WriteEntry -Entry $Child -IndentLevel $IndentLevel
            }
        }

        if (-not $HideProgress) {
            Write-Progress -Activity "Generating Excel sheet..." -Status "Almost done..." -PercentComplete 100
        }

        # helper function to set the maximum size of a row or column
        function ClampSize
        {
            param (
                $Range,
                $MaxWidth = 0,
                $MaxHeight = 0
            )

            if ($MaxWidth -gt 0)
            {
                if ($Range.ColumnWidth -gt $MaxWidth)
                {
                    $Range.ColumnWidth = $MaxWidth
                }
            }
            if ($MaxHeight -gt 0)
            {
                if ($Range.RowHeight -gt $MaxHeight)
                {
                    $Range.RowHeight = $MaxHeight
                }
            }
        }

        # set column sizes
        $ws.Columns("A:F").ColumnWidth = 70
        $ws.Columns.AutoFit() | Out-Null
        $ws.Columns("C").ColumnWidth = 70
        $ws.Columns("E").ColumnWidth = 8.43
        ClampSize -Range $ws.Columns("F") -MaxWidth 100

        for ($i = 3; $i -le $CurrentRow; $i++) {
            ClampSize -Range $ws.Rows("$i") -MaxHeight 40
        }

        # apply gray borders
        $ws.Range("A2:F$CurrentRow").Borders.Color = 0x808080
        $ws.Range("A2:F$CurrentRow").Borders.LineStyle = 1

        # freeze top row
        $ws.Rows("3").Select() | Out-Null
        $excel.ActiveWindow.FreezePanes = $true
        $ws.Range("A1").Select() | Out-Null

        # save
        if ($null -ne $ExportPath)
        {
            if ($ExportPath.Extension -eq ".xlsx") {
                $ws.SaveAs($ExportPath.FullName)
            } else {
                $ws.SaveAs($ExportPath.FullName, 52)
            }
        }

        # show excel
        $excel.Visible = $Show
        $excel.DisplayAlerts = $true
    }
    catch
    {
        Write-Error $_
    }
    finally
    {
        if ($excel.Visible -eq $false)
        {
            if ($null -ne $wb) { $wb.Close() }
            if ($null -ne $excel) { $excel.Quit() }
        }

        if ($null -ne $ws) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null }
        if ($null -ne $wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
        if ($null -ne $excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
        [GC]::Collect()
    }
    
    Write-Progress -Activity "Generating Excel sheet..." -Completed
}