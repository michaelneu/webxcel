. .\variables.ps1
. .\constants.ps1
. .\log.ps1

LogInfo "Collecting modules"

$missing = [System.Reflection.Missing]::Value
$excel = New-Object -ComObject Excel.Application
$book = $excel.Workbooks.Open($FILENAME, $missing, $true)
$modules = $book.VBProject.VBComponents;
$exportedModules = 0

For ($moduleIndex = 0; $moduleIndex -lt $modules.Count; $moduleIndex++)
{
    $module = $modules.Item($moduleIndex + 1)
    $moduleFilename = switch ($module.Type)
    {
        $COMPONENT_TYPE_MODULE { "src\Modules\$($module.Name).bas" }
        $COMPONENT_TYPE_CLASS { "src\Classes\$($module.Name).cls" }
        default { "" }
    }

    if ($moduleFilename -eq "")
    {
        echo "skipping module '$($module.Name)'"
        continue
    }

    $moduleDestination = [IO.Path]::Combine($CWD, $moduleFilename)
    echo "exporting $moduleFilename"
    $module.Export($moduleDestination)
    $exportedModules += 1
}

$excel.Quit()
LogInfo "Exported $exportedModules modules"
LogEmptyLine
