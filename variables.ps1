$CWD = (Resolve-Path .\).Path
$BUILD_DIRECTORY = [IO.Path]::Combine($CWD, "build")
$FILENAME = [IO.Path]::Combine($BUILD_DIRECTORY, "webxcel.xlsm")
