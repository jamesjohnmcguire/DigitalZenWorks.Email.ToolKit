REM %1 - Type of build
REM %2 - Version (such as 1.0.0.5)
REM %3 - API key

CD %~dp0
CD ..

IF "%1"=="publish" GOTO publish
IF "%1"=="release" GOTO release

:default
msbuild -property:Configuration=Release;IncludeAllContentForSelfExtract=true;OutputPath=Bin\Release\;Platform="Any CPU";PublishReadyToRun=true;PublishSingleFile=true;Runtimeidentifier=win-x64;SelfContained=true;TargetFramework=net6.0-windows -restore -target:publish;rebuild ToolKit.Application

GOTO end

:publish

msbuild -property:Configuration=Release;OutputPath=Bin\Release\Library;TargetFramework=net6.0-windows -restore -target:rebuild ToolKit.Library
msbuild -property:Configuration=Release;OutputPath=Bin\Release\Library -restore -target:pack ToolKit.Library
cd ToolKit.Library\Bin\Release\Library
nuget push DigitalZenWorks.Email.ToolKit.%2.nupkg %3 -Source https://api.nuget.org/v3/index.json
GOTO end

:release

msbuild -property:Configuration=Release;IncludeAllContentForSelfExtract=true;OutputPath=Bin\;Platform="Any CPU";PublishReadyToRun=true;PublishSingleFile=true;Runtimeidentifier=win-x64;SelfContained=true;TargetFramework=net6.0-windows -restore -target:publish;rebuild ToolKit.Application
CD ToolKit.Application\Bin\publish
7z u DigitalZenWorks.Email.ToolKit.zip

gh release create v%2 --notes "%2" DigitalZenWorks.Email.ToolKit.zip

CD ..\..\..
:end
