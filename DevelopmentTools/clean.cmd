::@ECHO OFF
CD %~dp0
CD ..

DEL /S /Q *.csproj.user

IF EXIST .vs\NUL RD /S /Q .vs
IF EXIST 1\NUL RD /S /Q 1
IF EXIST Bin\NUL RD /S /Q Bin

IF EXIST TestConsoleApp\bin\NUL RD /S /Q TestConsoleApp\bin
IF EXIST TestConsoleApp\obj\NUL RD /S /Q TestConsoleApp\obj
IF EXIST ToolKit.Application\Bin\NUL RD /S /Q ToolKit.Application\Bin
IF EXIST ToolKit.Application\obj\NUL RD /S /Q ToolKit.Application\obj
IF EXIST ToolKit.Library\Bin\NUL RD /S /Q ToolKit.Library\Bin
IF EXIST ToolKit.Library\obj\NUL RD /S /Q ToolKit.Library\obj
IF EXIST ToolKit.Tests\bin\NUL RD /S /Q ToolKit.Tests\bin
IF EXIST ToolKit.Tests\obj\NUL RD /S /Q ToolKit.Tests\obj
IF EXIST CommandLineCommands\CommandLineCommands.Tests\bin\NUL RD /S /Q CommandLineCommands\CommandLineCommands.Tests\bin
IF EXIST CommandLineCommands\CommandLineCommands.Tests\obj\NUL RD /S /Q CommandLineCommands\CommandLineCommands.Tests\obj
IF EXIST CommandLineCommands\CommandLineCommands\bin\NUL RD /S /Q CommandLineCommands\CommandLineCommands\bin
IF EXIST CommandLineCommands\CommandLineCommands\obj\NUL RD /S /Q CommandLineCommands\CommandLineCommands\obj
IF EXIST DbxOutlookExpress\DbxOutlookExpress\1\NUL RD /S /Q DbxOutlookExpress\DbxOutlookExpress\1
IF EXIST DbxOutlookExpress\DbxOutlookExpress\bin\NUL RD /S /Q DbxOutlookExpress\DbxOutlookExpress\bin
IF EXIST DbxOutlookExpress\DbxOutlookExpress\obj\NUL RD /S /Q DbxOutlookExpress\DbxOutlookExpress\obj
IF EXIST DbxOutlookExpress\DbxOutlookExpressTests\bin\NUL RD /S /Q DbxOutlookExpress\DbxOutlookExpressTests\bin
IF EXIST DbxOutlookExpress\DbxOutlookExpressTests\obj\NUL RD /S /Q DbxOutlookExpress\DbxOutlookExpressTests\obj
IF EXIST MsgKit\MsgKit\Bin\NUL RD /S /Q MsgKit\MsgKit\Bin
IF EXIST MsgKit\MsgKit\obj\NUL RD /S /Q MsgKit\MsgKit\obj
IF EXIST UtilitiesNet\UtilitiesNetLibrary\bin\NUL RD /S /Q UtilitiesNet\UtilitiesNetLibrary\bin
IF EXIST UtilitiesNet\UtilitiesNetLibrary\obj\NUL RD /S /Q UtilitiesNet\UtilitiesNetLibrary\obj
IF EXIST UtilitiesNet\UtilitiesNetTests\bin\NUL RD /S /Q UtilitiesNet\UtilitiesNetTests\bin
IF EXIST UtilitiesNet\UtilitiesNetTests\obj\NUL RD /S /Q UtilitiesNet\UtilitiesNetTests\obj
