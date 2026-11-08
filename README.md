# DigitalZenWorks.Email.ToolKit

[![License](https://img.shields.io/github/license/jamesjohnmcguire/DigitalZenWorks.Email.ToolKit)](LICENSE)
[![NuGet Publish](https://github.com/jamesjohnmcguire/DigitalZenWorks.Email.ToolKit/actions/workflows/build.yml/badge.svg?branch=main&event=workflow_dispatch)](https://github.com/jamesjohnmcguire/DigitalZenWorks.Email.ToolKit/actions/workflows/build.yml)
[![Latest Release](https://img.shields.io/github/v/release/jamesjohnmcguire/DigitalZenWorks.Email.ToolKit)](../../releases)
[![Last Commit](https://img.shields.io/github/last-commit/jamesjohnmcguire/DigitalZenWorks.Email.ToolKit)](../../commits)
[![Open Issues](https://img.shields.io/github/issues/jamesjohnmcguire/DigitalZenWorks.Email.ToolKit)](../../issues)
[![.NET](https://img.shields.io/badge/.NET-blueviolet)](https://dotnet.microsoft.com/)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen)](#contributing)

This is a C# application and library for interacting with email messages and storage.

It is a useful utility for converting email storage from Outlook Express, Windows Live Mail and Windows Mail.  It can migrate all .dbx and .eml files to .pst (Outlook) files.

## Documentation

| Purpose | Start here |
| --- | --- |
| Using the NuGet library | [API overview](#library-and-api-usage) |
| Connection behavior | [Outlook lifecycle](Documentation/OutlookLifecycle.md) |
| Maintaining this repository | [Maintainer notes][maintainer-notes] |
| Release changes | [Release notes](Documentation/ReleaseNotes.txt) |

[maintainer-notes]: Documentation/DevelopmentNotes.md

## Installation

### Prerequisites

The Windows application requires a working Microsoft Outlook installation
for Outlook/PST operations. `--help` does not connect to Outlook.

### Download

Download the Windows x64 ZIP from the Releases page. Extract `Det.exe` and
place it in a directory on your PATH.


### Nuget
PM> Install-Package DigitalZenWorks.Email.ToolKit

### Build from source

This project includes the [DbxOutlookExpress project](https://github.com/jamesjohnmcguire/DbxOutlookExpress) as a submodule.  So, be sure to include submodules when retreiving the repository contents.

git clone --recurse-submodules https://github.com/jamesjohnmcguire/DigitalZenWorks.Email.ToolKit

See [Maintainer notes](Documentation/DevelopmentNotes.md) for local builds,
release preparation, and explicit publishing commands.

### Outlook service lifetime

Reuse one `OutlookService` on the same STA thread for a workflow. Repeated
successful `Connect()` calls reuse its connection and session.
`Disconnect()` detaches without quitting a potentially shared Outlook.
The timeout bounds the requested activation-probe wait, including cleanup;
it neither cancels COM activation nor limits the whole `Connect()` call.

See [Outlook lifetime and timeout design](Documentation/OutlookLifecycle.md)
for ownership, retry behavior, compatibility, and the reasons for keeping
the temporary STA probe separate from the real connection.

### Usage:

NOTE: Always back up any data you might be modifying.  This package has been tested with hundreds of thousands of messages,
more than 50 GBs of data, without any problems.  But there can always be a first.  So, please back up your data before using this tool.  Did I mention that you should back up?

This is, currently, a command line program to be run from the command line or terminal.

#### Command line usage:

Det.exe \<command\> \<source-path\> \<destination-path\>

| Commands:             |                                 | options        | parameters                                         |
| --------------------- | ------------------------------- | -------------- | -------------------------------------------------- |
| add-pst               | Add given PST to the profile    |                | PST file path                                      |
| dbx-to-pst            | Migrate dbx files to pst file   | -e, --encoding | dbx files path, PST file path                      |
| details               | Show details of given item      |                | PST file path, entry id                            |
| eml-to-pst            | Migrate eml files to pst file   | -a, --adjust   | eml files path, PST file path                      |
| list-folders          | List all sub folders of a given | -r, --recurse  | PST file path, folder path                         |
|                       | folder                          |                |                                                    |
| list-ids              | List all entry IDs of items in  |                | PST file path, folder path                         |
|                       | a given folder                  |                |                                                    |
| list-top-senders      | List the top senders of a given | -c, --count    | PST file path                                      |
|                       | store                           |                |                                                    |
| list-total-duplicates | List all duplicates in a given  |                | PST file path                                      |
|                       | store                           |                |                                                    |
| merge-folders         | Merge duplicate Outlook folders | -n, --dryrun   |                                                    |
| merge-stores          | Merge one store into another    |                | source PST file path, destination PST file path    |
| move-folder           | Move one folder to another      |                | source PST file path, source folder path           |
|                       |                                 |                | destination PST file path, destination folder path |
| remove-duplicates     | Remove duplicate Outlook items  | -n, --dryrun   | PST file path                                      |
|                       |                                 | -s, --flush    |                                                    |
| remove-empty-folders  | Prune empty folders             |                | PST file path                                      |
| help                  | Display this information        |                |                                                    |

&nbsp;
&nbsp;
##### Command line usage notes:
The command is optional if the command can be inferred from the source-path.  For example, if the source path is a directory containing *.eml files, they will processed accordingly.
If the source-path is a directory, the command will attempt to process the files in directory.  If the source-path is a file, it will process that file directly.
Destination PST paths may omit `.pst`; creation appends the suffix
(for example, `archive` creates `archive.pst`).

###### Outlook connection timeout

All Outlook commands accept `--timeout <seconds>` or `-t <seconds>`. The
value must be an integer from `0` to `2147483`; the default is `10` seconds.
Invalid values fail before an Outlook connection is attempted. A value of
`0` starts or joins the activation probe without waiting for it to finish.

This option bounds the activation-probe wait, including cleanup. It does
not cancel COM activation or set a deadline for subsequent connection,
migration, or other Outlook operations. Help does not connect to Outlook.

```shell
Det.exe merge-stores --timeout 30 source.pst destination.pst
Det.exe list-folders -t 30 archive.pst
```

###### Examples
Det.exe remove-duplicates --dryrun \path\to\some.pst
Det.exe dbx-to-pst --encoding shift_jis \path\to\some-dbx-files
Det.exe dbx-to-pst \path\to\some-dbx-files \path\to\some.pst
Det.exe eml-to-pst --adjust "%USERPROFILE%\AppData\Local\Microsoft\Windows Live Mail\Storage Folders" %USERPROFILE%\Import.pst
Det.exe move-folder \path\to\some.pst source\folder\path \path\to\some.pst destination\folder\path
Det.exe list-folders \path\to\some.pst some\folder\path
Det.exe list-top-senders \path\to\some.pst
Det.exe list-total-duplicates \path\to\some.pst

###### dbx-to-pst
The optional --encoding option allows you to add a preferred encoding, in the rare case, that encoding can not be detected properly.  It must be a string that is recognized by Encoding.GetEncoding, along with Encoding.RegisterProvider(CodePagesEncodingProvider). Valid encoding strings can be found here: https://learn.microsoft.com/en-us/dotnet/api/system.text.encoding?view=net-7.0
If a PST file location is not specified, it will use the base location of the DBX files.

###### eml-to-pst
Some applications, like 'Windows Live Mail' often include interim folders, such as 'Imported Folders', 'Local Folders' and 'Storage Folders', creating unnecessary levels of nested folders.  Using the --adjust option will remove these interim folders in the PST file.

###### merge-folders
Sometimes, folders get duplicated, like in the following manner:
Testing
Testing (1)
Testing (1) (1)
Testing (1) (2)

If you ever try to move or copy a folder to a place where a folder with that name exists, Outlook will add, but will give it a name an appendix like ' (1)'.  In import or export processes, often these are the exact same folders, so you can end up with multiple duplicate folders like this.  This will merge these folders into a single folder.  If there are duplicate mail items, these will copied.  So, this wil not remove the duplicate mail items.  But, it doesn't create any duplicates and the merging of folders, is an essential precursor to the eventual duplicates removal.

###### move-folder
The third and fifth parameters specify the folder path with Outlook itself, starting with the root folder name.  The root folder name will almost always be the same as the PST file name.  You can also check at the very top folder for the PST within Outlook.

###### remove-duplicates

Removes duplicate Outlook items. If no PST path is provided, it attempts to
remove duplicates across all folders and stores in the current Outlook
account.

Use `--dryrun` to preview which items would be deleted. Use `--flush` to empty
the Deleted Items folder afterwards. For large stores, this can take time.

## Library and API usage

For new Outlook functionality, start with `OutlookService` through
`IOutlookService`. A successful connection provides a stable
`IOutlookSession` for the workflow. Store and folder helpers perform more
specific operations using that session.

| API | Responsibility |
| --- | --- |
| `OutlookService` / `IOutlookService` | Connect, expose a session, disconnect |
| `OutlookSession` / `IOutlookSession` | Wrap Outlook session operations |
| `OutlookStore` | Work with PST files and stores |
| `OutlookFolder` | Work with folders and their items |
| `OutlookItem` | Inspect and compare message content |
| `Migrate` | Import DBX and EML messages into PST files |
| `OutlookAccount` | Legacy singleton retained for existing clients |

This synchronous console example establishes the required STA context and
shows how the service, session, and store helper fit together:

```csharp
namespace EmailToolKitExample;

using System;
using DigitalZenWorks.Email.ToolKit;

internal static class Program
{
	[STAThread]
	private static int Main()
	{
		IOutlookService outlook = new OutlookService();

		try
		{
			bool connected = outlook.Connect(timeOutSeconds: 30);

			if (connected == false)
			{
				Console.Error.WriteLine("Outlook is unavailable.");
				return 1;
			}

			IOutlookSession? session = outlook.Session;

			if (session == null)
			{
				return 1;
			}

			OutlookStore stores = OutlookStore.FromSession(session);
			// Use stores and session for this workflow's operations.
			return 0;
		}
		finally
		{
			outlook.Disconnect();
		}
	}
}
```

Keep the workflow on the same STA thread. Normal callers do not need to
create factories or connection objects. Create helpers with
`OutlookStore.FromSession(session)` or `OutlookFolder.FromSession(session)`.
Both currently require the concrete `OutlookSession` supplied by the
service; custom session implementations are not yet supported.
See [Outlook lifecycle][lifecycle]
for the timeout boundary, ownership rules, and connection diagnostics.
The NuGet package includes XML documentation for member details through
IntelliSense and IDE API browsing.

### Existing migration APIs

The static migration entry points remain available:

```csharp
Migrate.DbxToPst(dbxPath, pstPath);
Migrate.EmlToPst(emlPath, pstPath);
```

`Migrate` currently obtains Outlook through the legacy
`OutlookAccount.Instance` infrastructure. Its connection is independent of
a separately created `OutlookService`. The singleton also remains available
to existing clients; new Outlook functionality should prefer the service
and session interfaces.

Read-receipt fields are preserved as ordinary attachments, defaulting to
`read-receipt.txt`. Conversion leaves the archived EML unchanged and does not
generate Content-IDs or infer missing HTML relationships. This preserves the
notification data without creating a native Outlook read-receipt object.

[lifecycle]: Documentation/OutlookLifecycle.md

## Contributing

If you have found a bug or have a suggestion that would make this better, please fork this repository and create a pull request. You can also simply open an issue with the tag "bug" or "enhancement".

### Process:

1. Fork the Project
2. Create your Bug / Feature Branch
3. Commit your Changes
4. Push to the Branch
5. Open a Pull Request - Explain the why the request and how the changes are beneficial. Explicitly mention the unit tests used in supporting this change.

### Coding style
Please match the current coding style.  Most notably:
1. One operation per line
2. Use complete English words in variable and method names
3. Attempt to declare variable and method names in a self-documenting manner
4. Use 4 column tabs and Allman style braces


## License

Distributed under the MIT License. See `LICENSE` for more information.

## Contact

James John McGuire - jamesjohnmcguire@gmail.com [LinkedIn](https://twitter.com/https://www.linkedin.com/in/jamesjohnmcguire) -  [GitHub](https://github.com/jamesjohnmcguire)

Project Link: [Email.ToolKit](https://github.com/jamesjohnmcguire/DigitalZenWorks.Email.ToolKit)
