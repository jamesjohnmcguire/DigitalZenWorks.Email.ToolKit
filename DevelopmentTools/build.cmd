REM %1 - Type of build
REM %2 - Version (such as 1.0.0.5)
REM %3 - API key

CD %~dp0
CD ..

IF "%1"=="publish" GOTO publish
IF "%1"=="release" GOTO release

:default
CD MsgKit
git checkout for-exe
CD ..

msbuild -property:Configuration=Release;IncludeAllContentForSelfExtract=true;OutputPath=Bin\;Platform="Any CPU";PublishReadyToRun=true;PublishSingleFile=true;Runtimeidentifier=win-x64;SelfContained=true;TargetFramework=net7.0-windows -restore -target:publish;rebuild ToolKit.Application

CD MsgKit
git checkout dzw-complete
CD ..

GOTO end

:publish

if "%~2"=="" GOTO error1
if "%~3"=="" GOTO error2

msbuild -property:Configuration=Release;OutputPath=Bin\Release\Library -restore -target:rebuild;pack ToolKit.Library

CD ToolKit.Library\Bin\Release\Library
nuget push DigitalZenWorks.Email.ToolKit.%2.nupkg %3 -Source https://api.nuget.org/v3/index.json
GOTO end

:release

msbuild -property:Configuration=Release;IncludeAllContentForSelfExtract=true;OutputPath=Bin\;Platform="Any CPU";PublishReadyToRun=true;PublishSingleFile=true;Runtimeidentifier=win-x64;SelfContained=true;TargetFramework=net7.0-windows -restore -target:publish;rebuild ToolKit.Application

CD ToolKit.Application\Bin\publish
7z u DigitalZenWorks.Email.ToolKit.zip

gh release create v%2 --notes "%2" DigitalZenWorks.Email.ToolKit.zip

CD ..\..\..

:error1
ECHO No version tag specified
GOTO end

:error2
ECHO No API key specified

:end
