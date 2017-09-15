$XLSM_FILE_NAME = "webxcel.xlsm"

$COMPONENT_TYPE_MODULE = 1
$COMPONENT_TYPE_CLASS = 2
$XL_FILE_FORMAT_MACRO_ENABLED = 52

$missing = [Reflection.Missing]::Value

$excel = New-Object -ComObject Excel.Application
$book = $excel.Workbooks.Add($missing)

Function AddScriptToBook($book, $file)
{
    $extension = $file.Extension.ToLower()

    $lines = [IO.File]::ReadAllLines($file.FullName)
    $lines = [Linq.Enumerable]::SkipWhile($lines, [Func[string, bool]]{ param($x)       $x.StartsWith("VERSION") `
                                                                                    -or $x.StartsWith("BEGIN") `
                                                                                    -or $x.StartsWith("  ") `
                                                                                    -or $x.StartsWith("END") `
                                                                                    -or $x.StartsWith("Attribute") `
                                })

    $code = [String]::Join("`r`n", $lines)

    $moduleType = $COMPONENT_TYPE_MODULE

    If ($extension -eq ".cls")
    {
        $moduleType = $COMPONENT_TYPE_CLASS
    }

    $module = $book.VBProject.VBComponents.Add($moduleType)

    $module.CodeModule.AddFromString($code)
    $module.Name = [IO.Path]::GetFileNameWithoutExtension($file.FullName)
}

$files = gci src *.* -rec | where { ! $_.PSIsContainer }

ForEach ($file in $files)
{
    AddScriptToBook $book $file
}

$build = Resolve-Path "build"
MkDir -Force $build > $null

$filename = [IO.Path]::Combine($build, $XLSM_FILE_NAME)

$excel.DisplayAlerts = $false
$book.SaveAs($filename, $XL_FILE_FORMAT_MACRO_ENABLED)
$excel.DisplayAlerts = $true

$excel.Quit()