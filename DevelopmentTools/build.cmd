@ECHO OFF
SETLOCAL

SET "toolkitBuildExitCode=0"
SET "toolkitAction=%~1"
SET "toolkitLibrary=SourceCode\ToolKit.Library\ToolKit.Library.csproj"
SET "toolkitApplication=SourceCode\ToolKit.Application\ToolKit.Application.csproj"
SET "toolkitRepository=jamesjohnmcguire/DigitalZenWorks.Email.ToolKit"

:: release uploads the executable; publish packs and publishes the library.
PUSHD "%~dp0.."
IF ERRORLEVEL 1 EXIT /B 1

IF "%~1"=="" GOTO debug
IF /I "%~1"=="debug" GOTO debug
IF /I "%~1"=="release" GOTO prepare
IF /I "%~1"=="publish" GOTO prepare

ECHO Usage: build.cmd [debug^|release^|publish] [expected-version]
SET "toolkitBuildExitCode=1"
GOTO end

:debug
CALL :require msbuild
IF ERRORLEVEL 1 GOTO end

msbuild SourceCode\DigitalZenWorks.Email.ToolKit.slnx -property:Configuration=Debug;OutputPath=Bin\;Platform="Any CPU" -restore -target:rebuild
SET "toolkitBuildExitCode=%ERRORLEVEL%"
GOTO end

:prepare
CALL :require dotnet
IF ERRORLEVEL 1 GOTO end
CALL :require git
IF ERRORLEVEL 1 GOTO end

CALL :readVersion "%toolkitLibrary%" toolkitVersion
IF ERRORLEVEL 1 GOTO end
CALL :readVersion "%toolkitApplication%" toolkitApplicationVersion
IF ERRORLEVEL 1 GOTO end
IF NOT "%toolkitVersion%"=="%toolkitApplicationVersion%" SET "message=Application and library versions must match."
IF NOT "%toolkitVersion%"=="%toolkitApplicationVersion%" GOTO error

IF NOT "%~2"=="" IF NOT "%~2"=="%toolkitVersion%" SET "message=Requested version differs from project version %toolkitVersion%."
IF NOT "%~2"=="" IF NOT "%~2"=="%toolkitVersion%" GOTO error

SET "toolkitCommit="
FOR /F "delims=" %%C IN ('git rev-parse HEAD') DO SET "toolkitCommit=%%C"
IF NOT DEFINED toolkitCommit SET "message=Unable to read the current Git commit."
IF NOT DEFINED toolkitCommit GOTO error

SET "toolkitArtifacts=%CD%\artifacts\%toolkitVersion%"
SET "toolkitArchive=%toolkitArtifacts%\DigitalZenWorks.Email.ToolKit.%toolkitVersion%.win-x64.zip"
SET "toolkitPackage=%toolkitArtifacts%\DigitalZenWorks.Email.ToolKit.%toolkitVersion%.nupkg"
CALL :require msbuild
IF ERRORLEVEL 1 GOTO end
CALL :require gh
IF ERRORLEVEL 1 GOTO end
IF NOT "%~3"=="" SET "message=API-key arguments are no longer supported. Configure NuGet trusted publishing."
IF NOT "%~3"=="" GOTO error
CALL :publicationCheck
IF ERRORLEVEL 1 GOTO end
IF NOT EXIST "%toolkitArtifacts%\." MKDIR "%toolkitArtifacts%"
IF NOT EXIST "%toolkitArtifacts%\." SET "message=Unable to create the artifacts directory."
IF NOT EXIST "%toolkitArtifacts%\." GOTO error
IF /I "%toolkitAction%"=="publish" GOTO publish

:release
CALL :require zip.exe
IF ERRORLEVEL 1 GOTO end

:staging
SET "toolkitStaging=%toolkitArtifacts%\publish-%RANDOM%-%RANDOM%"
IF EXIST "%toolkitStaging%" GOTO staging
IF EXIST "%toolkitStaging%.zip" GOTO staging

msbuild "%toolkitApplication%" -restore -target:Publish -property:Configuration=Release;DebugType=embedded;IncludeAllContentForSelfExtract=true;PublishReadyToRun=true;PublishSingleFile=true;RuntimeIdentifier=win-x64;SelfContained=true;GeneratePackageOnBuild=false "-property:PublishDir=%toolkitStaging%/"

SET "toolkitBuildExitCode=%ERRORLEVEL%"
IF NOT "%toolkitBuildExitCode%"=="0" GOTO end
IF NOT EXIST "%toolkitStaging%\Det.exe" SET "message=Publish succeeded but Det.exe was not produced."
IF NOT EXIST "%toolkitStaging%\Det.exe" GOTO error

PUSHD "%toolkitStaging%"
IF ERRORLEVEL 1 SET "message=Unable to enter the publish directory."
IF ERRORLEVEL 1 GOTO error
zip.exe -r "%toolkitStaging%.zip" .
SET "toolkitBuildExitCode=%ERRORLEVEL%"
POPD
IF NOT "%toolkitBuildExitCode%"=="0" GOTO end

MOVE /Y "%toolkitStaging%.zip" "%toolkitArchive%" >NUL
SET "toolkitBuildExitCode=%ERRORLEVEL%"
IF NOT "%toolkitBuildExitCode%"=="0" GOTO end

ECHO Prepared executable archive: %toolkitArchive%
CALL :publicationCheck
IF ERRORLEVEL 1 GOTO end
CALL :uploadArtifact "%toolkitArchive%"
SET "toolkitBuildExitCode=%ERRORLEVEL%"
GOTO end

:publish
msbuild "%toolkitLibrary%" -restore -target:Pack -property:Configuration=Release;GeneratePackageOnBuild=false "-property:PackageOutputPath=%toolkitArtifacts%"
SET "toolkitBuildExitCode=%ERRORLEVEL%"
IF NOT "%toolkitBuildExitCode%"=="0" GOTO end

IF NOT EXIST "%toolkitPackage%" SET "message=Pack succeeded but the expected NuGet package was not produced."
IF NOT EXIST "%toolkitPackage%" GOTO error
ECHO Prepared NuGet package: %toolkitPackage%
CALL :publicationCheck
IF ERRORLEVEL 1 GOTO end
CALL :uploadArtifact "%toolkitPackage%"
IF ERRORLEVEL 1 GOTO end
gh workflow run build.yml --ref main --repo "%toolkitRepository%" -f version="%toolkitVersion%"
SET "toolkitBuildExitCode=%ERRORLEVEL%"
IF NOT "%toolkitBuildExitCode%"=="0" GOTO end
ECHO NuGet publication requested. Follow the workflow with:
ECHO gh run watch --repo "%toolkitRepository%" --exit-status
GOTO end

:uploadArtifact
gh release view "v%toolkitVersion%" --repo "%toolkitRepository%" >NUL 2>&1
IF ERRORLEVEL 1 GOTO createGitHubRelease
gh release upload "v%toolkitVersion%" "%~1" --clobber --repo "%toolkitRepository%"
SET "toolkitBuildExitCode=%ERRORLEVEL%"
IF NOT "%toolkitBuildExitCode%"=="0" EXIT /B %toolkitBuildExitCode%
gh release edit "v%toolkitVersion%" --draft=false --repo "%toolkitRepository%"
SET "toolkitBuildExitCode=%ERRORLEVEL%"
EXIT /B %toolkitBuildExitCode%

:createGitHubRelease
gh release create "v%toolkitVersion%" "%~1" --target "%toolkitCommit%" --notes-file Documentation\ReleaseNotes.txt --repo "%toolkitRepository%"
SET "toolkitBuildExitCode=%ERRORLEVEL%"
EXIT /B %toolkitBuildExitCode%

:require
WHERE "%~1" >NUL 2>&1
SET "toolkitBuildExitCode=%ERRORLEVEL%"
IF "%toolkitBuildExitCode%"=="0" EXIT /B 0
ECHO %~1 is required. For MSBuild, use a Visual Studio Developer Command Prompt.
EXIT /B %toolkitBuildExitCode%

:readVersion
SET "toolkitVersionFile=%TEMP%\EmailToolKit-version-%RANDOM%-%RANDOM%.txt"
dotnet msbuild "%~1" -nologo -getProperty:Version "-p:Configuration=Release" > "%toolkitVersionFile%"
SET "toolkitBuildExitCode=%ERRORLEVEL%"
SET "toolkitProjectVersion="
IF "%toolkitBuildExitCode%"=="0" SET /P "toolkitProjectVersion="<"%toolkitVersionFile%"
DEL /Q "%toolkitVersionFile%" >NUL 2>&1
IF NOT "%toolkitBuildExitCode%"=="0" EXIT /B %toolkitBuildExitCode%

IF NOT DEFINED toolkitProjectVersion ECHO Unable to read the project version.
IF NOT DEFINED toolkitProjectVersion SET "toolkitBuildExitCode=1"
IF NOT "%toolkitBuildExitCode%"=="0" EXIT /B %toolkitBuildExitCode%

SET "%~2=%toolkitProjectVersion%"
EXIT /B 0

:publicationCheck
SET "toolkitBranch="
FOR /F "delims=" %%B IN ('git branch --show-current') DO SET "toolkitBranch=%%B"
IF NOT "%toolkitBranch%"=="main" ECHO Release and publication require the main branch.
IF NOT "%toolkitBranch%"=="main" SET "toolkitBuildExitCode=1"
IF NOT "%toolkitBuildExitCode%"=="0" EXIT /B %toolkitBuildExitCode%
SET "toolkitDirty="
FOR /F "delims=" %%S IN ('git status --porcelain') DO SET "toolkitDirty=1"
IF DEFINED toolkitDirty ECHO Commit all changes, including submodules, before releasing or publishing.
IF DEFINED toolkitDirty SET "toolkitBuildExitCode=1"
IF NOT "%toolkitBuildExitCode%"=="0" EXIT /B %toolkitBuildExitCode%
SET "toolkitCurrentCommit="
FOR /F "delims=" %%C IN ('git rev-parse HEAD') DO SET "toolkitCurrentCommit=%%C"
IF NOT "%toolkitCurrentCommit%"=="%toolkitCommit%" ECHO The source commit changed during this command. Retry from a stable commit.
IF NOT "%toolkitCurrentCommit%"=="%toolkitCommit%" SET "toolkitBuildExitCode=1"
IF NOT "%toolkitBuildExitCode%"=="0" EXIT /B %toolkitBuildExitCode%
EXIT /B %toolkitBuildExitCode%

:error
ECHO %message%
SET "toolkitBuildExitCode=1"

:end
POPD
EXIT /B %toolkitBuildExitCode%
