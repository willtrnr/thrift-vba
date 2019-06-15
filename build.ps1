$ADODB_GUID = "{2A75196C-D9EB-4129-B803-931327F72D5C}"
$MSXML_GUID = "{F5078F18-C551-11D3-89B9-0000F81FE221}"
$SCRIPRUN_GUID = "{420B2830-E718-11CF-893D-00A0C9054228}"

$BUILD_FILENAME = "a.xlsm"

$LOCAL = Split-Path -Parent $MyInvocation.MyCommand.Definition

Add-Type -AssemblyName "Microsoft.Vbe.Interop"
Add-Type -AssemblyName "Microsoft.Office.Interop.Excel"

$missing = [Reflection.Missing]::Value

$excel = New-Object Microsoft.Office.Interop.Excel.ApplicationClass
$excel.DisplayAlerts = $false

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\Excel\Security" -Name "AccessVBOM" -PropertyType "DWORD" -Value 1 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\Excel\Security" -Name "VBAWarnings" -PropertyType "DWORD" -Value 1 -Force | Out-Null

$wb = $excel.Workbooks.Add($missing)
$vbproj = $wb.VBProject
$vbcomps = $vbproj.VBComponents

$vbproj.References.AddFromGuid($ADODB_GUID, 2, 8) | Out-Null
$vbproj.References.AddFromGuid($MSXML_GUID, 6, 0) | Out-Null
$vbproj.References.AddFromGuid($SCRIPRUN_GUID, 1, 0) | Out-Null

ForEach ($file in (Get-ChildItem -Path ([IO.Path]::Combine($LOCAL, "src")))) {
    $ext = $file.Extension.ToLower()
    $name = [IO.Path]::GetFileNameWithoutExtension($file.FullName)

    $module = $null
    If ($ext -eq ".cls") {
        $module = $vbcomps.Add([Microsoft.Vbe.Interop.vbextFileTypes]::vbextFileTypeClass)
    } ElseIf ($ext -eq ".bas") {
        $module = $vbcomps.Add([Microsoft.Vbe.Interop.vbextFileTypes]::vbextFileTypeModule)
    }

    If ($module) {
        $module.Name = $name
        $module.CodeModule.DeleteLines(1, $module.CodeModule.CountOfLines)
        $module.CodeModule.AddFromFile($file.FullName)
        If ($module.Type -eq [Microsoft.Vbe.Interop.vbextFileTypes]::vbextFileTypeClass) {
            $module.CodeModule.DeleteLines(1, 4)
        }
    }
}

$wb.SaveAs(([IO.Path]::Combine($LOCAL, $BUILD_FILENAME)), [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled, $missing, $missing, $missing, $missing, $missing, $missing, $false, $missing, $missing, $missing)

$wb.Close($false, $missing, $missing)
$wb = $null

$excel.Quit()
$excel = $null

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()