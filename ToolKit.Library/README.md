# DigitalZenWorks.Email.ToolKit

This is a C# library for interacting with email messages and storage.

### Usage:
The main classes are MapiItem, Migrate, OutlookOutlook, OutlookFolder and OutlookStore.  
Migrate is for migrating messages in other formats into Outlook.  Some examples:  
Migrate.DbxToPst(dbxPath, pstPath);  
Migrate.EmlToPst(dbxPath, pstPath);  

OutlookAccount acts on the entire default or current Outlook account.  OutlookStore acts on a specific store (PST file).  OutlookFolder acts on a specific folder. 
For cleaning up Outlook, you could use the following:  
OutlookAccount outlookAccont = new ();  
outlookAccont.MergeFolders();  
outlookAccont.RemoveDuplicates();  

For a more in-depth example and additional documentation, please refer to the application part of the project at [DigitalZenWorks.Email.ToolKit Project](https://github.com/jamesjohnmcguire/.Email.ToolKit)  

## Contributing

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".

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
