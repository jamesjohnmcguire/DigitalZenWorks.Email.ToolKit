::@ECHO OFF
CD %~dp0
CD ..

RD /S /Q .vs
RD /S /Q 1
RD /S /Q Bin
RD /S /Q TestConsoleApp\obj
RD /S /Q ToolKit.Application\Bin
RD /S /Q ToolKit.Application\obj
RD /S /Q ToolKit.Library\Bin
RD /S /Q ToolKit.Library\obj
RD /S /Q ToolKit.Tests\bin
RD /S /Q ToolKit.Tests\obj
RD /S /Q CommandLineCommands\CommandLineCommands.Tests\obj
RD /S /Q CommandLineCommands\CommandLineCommands\obj
RD /S /Q DbxOutlookExpress\DbxOutlookExpress\1
RD /S /Q DbxOutlookExpress\DbxOutlookExpress\obj
RD /S /Q DbxOutlookExpress\DbxOutlookExpressTests\bin
RD /S /Q DbxOutlookExpress\DbxOutlookExpressTests\obj
RD /S /Q MsgKit\MsgKit\Bin
RD /S /Q MsgKit\MsgKit\obj
RD /S /Q UtilitiesNet\UtilitiesNetLibrary\bin
RD /S /Q UtilitiesNet\UtilitiesNetLibrary\obj
RD /S /Q UtilitiesNet\UtilitiesNetTests\bin
RD /S /Q UtilitiesNet\UtilitiesNetTests\obj
