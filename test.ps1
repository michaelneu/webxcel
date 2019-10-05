. .\variables.ps1
. .\constants.ps1
. .\log.ps1

LogInfo "Collecting tests"

$missing = [System.Reflection.Missing]::Value
$excel = New-Object -ComObject Excel.Application
$book = $excel.Workbooks.Open($FILENAME, $missing, $true)
$modules = $book.VBProject.VBComponents;
$suites = @{}
$suiteFlags = @{}

$SUITE_FLAG_BEFORE_ALL = 1
$SUITE_FLAG_AFTER_ALL = 2
$SUITE_FLAG_BEFORE_EACH = 4
$SUITE_FLAG_AFTER_EACH = 8

$testCount = 0

For ($moduleIndex = 0; $moduleIndex -lt $modules.Count; $moduleIndex++)
{
    # vba interop seems to be written in VB6 => indices start at 1
    $module = $modules.Item($moduleIndex + 1)

    If (!$module.Name.StartsWith("Test"))
    {
        Continue
    }

    $code = $module.CodeModule.Lines(1, $module.CodeModule.CountOfLines)
    $lines = $code.Split("`r`n")

    $suiteFlags[$module.Name] = 0

    ForEach ($line in $lines)
    {
        If ($line.StartsWith("Public Function Test"))
        {
            If (!$suites.ContainsKey($module.Name))
            {
                $suites[$module.Name] = [System.Collections.ArrayList]@()
            }

            $testName = $line.Split(" ")[2].Trim("()")
            $_ = $suites[$module.Name].Add($testName)
            $testCount += 1
        }

        If ($line.StartsWith("Public Sub BeforeAll()"))
        {
            $suiteFlags[$module.Name] = $suiteFlags[$module.Name] -bor $SUITE_FLAG_BEFORE_ALL
        }

        If ($line.StartsWith("Public Sub AfterAll()"))
        {
            $suiteFlags[$module.Name] = $suiteFlags[$module.Name] -bor $SUITE_FLAG_AFTER_ALL
        }

        If ($line.StartsWith("Public Sub BeforeEach()"))
        {
            $suiteFlags[$module.Name] = $suiteFlags[$module.Name] -bor $SUITE_FLAG_BEFORE_EACH
        }

        If ($line.StartsWith("Public Sub AfterEach()"))
        {
            $suiteFlags[$module.Name] = $suiteFlags[$module.Name] -bor $SUITE_FLAG_AFTER_EACH
        }
    }
}

LogInfo "Found $($suites.Count) suites with $testCount tests"
LogEmptyLine

$passedSuites = 0
$passedTests = 0

Function HasSuiteFlag($flags, $flag)
{
    Return ($flags -band $flag) -eq $flag
}

ForEach ($suite in $suites.Keys)
{
    $successful = $true
    $tests = $suites[$suite]
    $flags = $suiteFlags[$suite]

    $title = $suite + " (" + $tests.Count + " tests)"
    echo $title

    $hasBeforeAll = HasSuiteFlag $flags $SUITE_FLAG_BEFORE_ALL
    $hasAfterAll = HasSuiteFlag $flags $SUITE_FLAG_AFTER_ALL
    $hasBeforeEach = HasSuiteFlag $flags $SUITE_FLAG_BEFORE_EACH
    $hasAfterEach = HasSuiteFlag $flags $SUITE_FLAG_AFTER_EACH

    If ($hasBeforeAll)
    {
        $excel.Run($suite + "." + "BeforeAll")
    }

    ForEach ($test in $tests)
    {
        If ($hasBeforeEach)
        {
            $excel.Run($suite + "." + "BeforeEach")
        }

        $result = $excel.Run($suite + "." + $test)

        Write-Host "  " -NoNewline

        if ($result.AssertSuccessful)
        {
            $passedTests += 1
            Write-Host " PASS " -BackgroundColor Green -ForegroundColor White -NoNewline
        }
        else
        {
            $successful = $false
            Write-Host " FAIL " -BackgroundColor Red -ForegroundColor White -NoNewline
        }

        Write-Host " $test" -NoNewline

        if ($result.AssertSuccessful)
        {
            Write-Host ": $($result.AssertMessage)" -ForegroundColor Gray
        }
        else
        {
            LogEmptyLine
            LogEmptyLine
            Write-Host $result.AssertMessage -ForegroundColor Red
            LogEmptyLine
        }

        If ($hasAfterEach)
        {
            $excel.Run($suite + "." + "AfterEach")
        }
    }

    If ($hasAfterAll)
    {
        $excel.Run($suite + "." + "AfterAll")
    }

    if ($successful)
    {
        $passedSuites += 1
    }

    echo ""
}

$excel.Quit()

Function LogSummary($title, $passed, $failed)
{
    Write-Host $title -NoNewline
    Write-Host "$passed passed" -ForegroundColor Green -NoNewline

    if ($failed -gt 0)
    {
        Write-Host ", " -NoNewline
        Write-Host "${failed} failed" -ForegroundColor Red -NoNewline
    }

    $total = $passed + $failed
    Write-Host ", $total total"
}

LogSummary "Test suites: " $passedSuites ($suites.Count - $passedSuites)
LogSummary "Tests:       " $passedTests ($testCount - $passedTests)
LogInfo "Ran all test suites."
LogEmptyLine
