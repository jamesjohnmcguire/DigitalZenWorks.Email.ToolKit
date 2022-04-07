REM %1 - Type of build
REM %2 - API key
REM %3 - Version (such as 1.0.0.5)

CD %~dp0
CD ..

IF "%1"=="publish" GOTO release
IF "%1"=="release" GOTO release

:default
msbuild -property:Configuration=Release;IncludeAllContentForSelfExtract=true;Platform="Any CPU";PublishReadyToRun=true;PublishSingleFile=true;Runtimeidentifier=win-x64;SelfContained=true;TargetFramework=net6.0-windows -restore -target:publish;rebuild ToolKit.Application

GOTO end

:publish

msbuild -property:Configuration=Release;OutputPath=Bin\Release\Library -restore -target:rebuild ToolKit.Library
msbuild -property:Configuration=Release;OutputPath=Bin\Release\Library -target:pack ToolKit.Library

cd ToolKit.Library\Bin\Release\Library

nuget push DigitalZenWorks.Email.ToolKit.%2.nupkg %3
GOTO end

:release

msbuild -property:Configuration=Release;IncludeAllContentForSelfExtract=true;OutputPath=Bin\;Platform="Any CPU";PublishReadyToRun=true;PublishSingleFile=true;Runtimeidentifier=win-x64;SelfContained=true;TargetFramework=net6.0-windows -restore -target:publish;rebuild ToolKit.Application
CD ToolKit.Application\Bin\publish
PAUSE
7z u DigitalZenWorks.Email.ToolKit.zip

gh release create v%2 --notes "%2" DigitalZenWorks.Email.ToolKit.zip

CD ..\..\..
:end
