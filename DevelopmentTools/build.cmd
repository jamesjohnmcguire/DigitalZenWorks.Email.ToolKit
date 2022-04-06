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

msbuild -property:Configuration=Release;OutputPath=Bin\Publish -restore -target:rebuild ToolKit.Library
msbuild -property:Configuration=Release;OutputPath=Bin\Publish -target:pack ToolKit.Library

cd ToolKit.Library\Bin\Publish

nuget push DigitalZenWorks.Email.ToolKit.%2.nupkg %3
GOTO end

:release

CALL msbuild -property:Configuration=Release;IncludeAllContentForSelfExtract=true;Platform="Any CPU";PublishReadyToRun=true;PublishSingleFile=true;Runtimeidentifier=win-x64;SelfContained=true;TargetFramework=net6.0-windows -restore -target:publish;rebuild ToolKit.Application
CD Bin\Release\AnyCpu\win-x64\publish

7z u DigitalZenWorks.Email.ToolKit.zip

gh release create v%2 --notes "%2" DigitalZenWorks.Email.ToolKit.zip

CD ..\..\..
:end
