. .\variables.ps1
. .\constants.ps1
. .\log.ps1
$missing = [Reflection.Missing]::Value

$excel = New-Object -ComObject Excel.Application
$book = $excel.Workbooks.Add($missing)

Function CleanupExportedCode($lines)
{
    While ($lines[0] -and ($lines[0].StartsWith("VERSION") -or $lines[0].StartsWith("BEGIN") -or $lines[0].StartsWith("  ") -or $lines[0].StartsWith("END")))
    {
        $ignore = $lines.RemoveAt(0)
    }
}

Function AddScriptToBook($book, $file)
{
    $extension = $file.Extension.ToLower()
    $lines = [System.Collections.ArrayList][IO.File]::ReadAllLines($file.FullName)
    CleanupExportedCode $lines

    $code = [String]::Join("`r`n", $lines.ToArray())
    $tmp = New-TemporaryFile
    [IO.File]::WriteAllLines($tmp.FullName, $code)

    $moduleType = $COMPONENT_TYPE_MODULE

    If ($extension -eq ".cls")
    {
        $moduleType = $COMPONENT_TYPE_CLASS
    }

    $module = $book.VBProject.VBComponents.Add($moduleType)
    $module.Name = [IO.Path]::GetFileNameWithoutExtension($file.FullName)
    $module.CodeModule.AddFromFile($tmp.FullName)
    Remove-Item $tmp.FullName -Force
}

$files = gci src *.* -rec | where { ! $_.PSIsContainer }
LogInfo "Found $($files.Count) modules"

ForEach ($file in $files)
{
    AddScriptToBook $book $file
}

MkDir -Force $BUILD_DIRECTORY > $null
$excel.DisplayAlerts = $false
$book.SaveAs($FILENAME, $XL_FILE_FORMAT_MACRO_ENABLED)
$excel.DisplayAlerts = $true

$excel.Quit()
LogInfo "Document created"
