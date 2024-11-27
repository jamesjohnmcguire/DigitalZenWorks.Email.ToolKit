REM %1 - Type of build
REM %2 - Version (such as 1.0.0.5)
REM %3 - API key

CD %~dp0
CD ..

IF "%1"=="publish" GOTO publish

:default
msbuild -property:Configuration=Release;IncludeAllContentForSelfExtract=true;OutputPath=Bin\;Platform="Any CPU";PublishReadyToRun=true;PublishSingleFile=true;Runtimeidentifier=win-x64;SelfContained=true -restore -target:publish;rebuild ToolKit.Application

IF "%1"=="release" GOTO release

GOTO end

:publish

if "%~2"=="" GOTO error1
if "%~3"=="" GOTO error2

CD ToolKit.Library

msbuild -property:Configuration=Release -restore -target:rebuild;pack ToolKit.Library.csproj

CD bin\Release
nuget push DigitalZenWorks.Email.ToolKit.%2.nupkg %3 -Source https://api.nuget.org/v3/index.json

CD ..\..\..

GOTO end

:release

CD ToolKit.Application\Bin\publish
7z u DigitalZenWorks.Email.ToolKit.zip
gh release create v%2 --notes-file ..\..\..\DevelopmentTools\ReleaseNotes.txt DigitalZenWorks.Email.ToolKit.zip

CD ..\..\..

GOTO end

:error1
ECHO No version tag specified
GOTO end

:error2
ECHO No API key specified

:end
