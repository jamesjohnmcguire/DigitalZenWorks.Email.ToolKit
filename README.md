# DigitalZenWorks.Email.ToolKit

This is a C# application and library for interacting with email messages and storage.

## Getting Started

### Prerequisites

This project includes the [DbxOutlookExpress project](https://github.com/jamesjohnmcguire/DbxOutlookExpress) as a submodule.  So, be sure to include submodules when retreiving the repository contents. 

### Installation
#### Git
git clone --recurse-submodules https://github.com/jamesjohnmcguire/DigitalZenWorks.Email.ToolKit

### Usage:

NOTE: Always back up any data you might be modifying.  So, please back up your data before using this tool.  Did I mention that you should back up?  

#### Command line usage:

DigitalZenWorks.Email.ToolKit \<command\> \<source-path\> \<destination-path\>

| Commands:            |                                 |
| -------------------- | ------------------------------  |
| dbx-to-pst           | Migrate dbx files to pst file   |
| eml-to-pst           | Migrate eml files to pst file   |
| merge-folders        | Merge duplicate Outlook folders |
| remove-empty-folders | Prune empty folders             |
| help                 | Display this information        |

##### Command line usage notes:
The command is optional if the command can be inferred from the source-path.  For example, if the source path is a directory containing *.eml files, they will processed accordingly.  
If the source-path is a directory, the command will attempt to process the files in directory.  If the source-path is a file, it will process that file directly.  
In regards to merge-folders, have you ever seen a folders like this:  
Testing  
Testing (1)  
Testing (1) (1)  
Testing (1) (2()  

If you ever try to move or copy a folder to a place where a folder with that name exists, Outlook will add, but will give it a name an appendix like ' (1)'.  In import or export processes, often these are the exact same folders, so you can end up with multiple duplicate folders like this.  This will merge these folders into a single folder.  If there are duplicate mail items, these will copied.  So, this wil not remove the duplicate mail items (That will come in the feature release).  But, it doesn't create any duplicates and the merging of folders, is an essential precursor to the eventual duplicates removal.  

## Contributing

Any contributions you make are **greatly appreciated**.  If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".

### Process:

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Coding style
Please match the current coding style.  Most notably:  
1. One operation per line
2. Use complete English words in variable and method names
3. Attempt to declare variable and method names in a self-documenting manner


## License

Distributed under the MIT License. See `LICENSE` for more information.

## Contact

James John McGuire - [@jamesmc](https://twitter.com/jamesmc) - jamesjohnmcguire@gmail.com

Project Link: [https://github.com/jamesjohnmcguire/DigitalZenWorks.Email.ToolKit](https://github.com/jamesjohnmcguire/DigitalZenWorks.Email.ToolKit)
