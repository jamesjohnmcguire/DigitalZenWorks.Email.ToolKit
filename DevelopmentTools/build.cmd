CD %~dp0
CD ..

IF "%1"=="release" GOTO release

SET NUGET_RESTORE_MSBUILD_ARGS=/p:Configuration=Release

CALL nuget restore DigitalZenWorks.Email.ToolKit.sln
CALL msbuild -property:Configuration=Release

GOTO end

:release
IF EXIST Bin\Release\AnyCPU\NUL DEL /Q Bin\Release\AnyCPU\*.*

CALL msbuild -property:Configuration=Release;TargetFramework=net6.0-windows -restore -target:publish
CD Bin\Release\AnyCpu

7z u DigitalZenWorks.Email.ToolKit.zip

gh release create v%1 --notes "%2" Bin\Release\AnyCPU\DigitalZenWorks.Email.ToolKit.zip

:end
