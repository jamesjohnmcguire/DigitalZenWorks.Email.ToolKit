CD %~dp0
CD ..

IF "%1"=="release" GOTO release

msbuild -property:Configuration=Release;IncludeAllContentForSelfExtract=true;Platform="Any CPU";PublishReadyToRun=true;PublishSingleFile=true;Runtimeidentifier=win-x64;SelfContained=true;TargetFramework=net6.0-windows -restore -target:publish;rebuild ToolKit.Application

GOTO end

:release
IF EXIST Bin\Release\AnyCPU\NUL DEL /Q Bin\Release\AnyCPU\*.*

CALL msbuild -property:Configuration=Release;IncludeAllContentForSelfExtract=true;Platform="Any CPU";PublishReadyToRun=true;PublishSingleFile=true;Runtimeidentifier=win-x64;SelfContained=true;TargetFramework=net6.0-windows -restore -target:publish;rebuild ToolKit.Application
CD Bin\Release\AnyCpu\win-x64\publish

7z u DigitalZenWorks.Email.ToolKit.zip

gh release create v%2 --notes "%2" DigitalZenWorks.Email.ToolKit.zip

CD ..\..\..
:end
