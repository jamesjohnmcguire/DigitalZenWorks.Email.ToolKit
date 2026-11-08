@ECHO OFF
SETLOCAL
SET "toolkitTestExitCode=0"

PUSHD "%~dp0.."
IF ERRORLEVEL 1 EXIT /B 1

WHERE msbuild >NUL 2>&1
IF ERRORLEVEL 1 ECHO Visual Studio MSBuild is required. Use a Visual Studio Developer Command Prompt.
IF ERRORLEVEL 1 SET "toolkitTestExitCode=1"
IF NOT "%toolkitTestExitCode%"=="0" GOTO end
WHERE dotnet >NUL 2>&1
IF ERRORLEVEL 1 ECHO The .NET SDK is required.
IF ERRORLEVEL 1 SET "toolkitTestExitCode=1"
IF NOT "%toolkitTestExitCode%"=="0" GOTO end

:: Build COM references with Visual Studio MSBuild, then run all solution tests.
msbuild SourceCode\DigitalZenWorks.Email.ToolKit.slnx -restore -target:Build -property:Configuration=Debug -maxcpucount -nologo
SET "toolkitTestExitCode=%ERRORLEVEL%"
IF NOT "%toolkitTestExitCode%"=="0" GOTO end

dotnet test SourceCode\DigitalZenWorks.Email.ToolKit.slnx --configuration Debug --no-build --no-restore --verbosity normal --logger trx --results-directory "%CD%\artifacts\TestResults" %*
SET "toolkitTestExitCode=%ERRORLEVEL%"

:end
POPD
EXIT /B %toolkitTestExitCode%
