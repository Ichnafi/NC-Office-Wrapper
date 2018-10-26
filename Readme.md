# Disclaimer
This is my first use of C# and Visual Studio ever. The code will be propably bad, there may be bugs, errors or just lack of knowledge but the code solved my problem. Who ever wants to make use of this, feel free to fork it, make it better, do what ever you want with it.
# Intention of this Programm
I had a problem with the behavior of Microsoft Office regarding lock-files (owner files). You can read the hole story here on [StackOverflow](https://stackoverflow.com/questions/50149662/ms-office-lock-file-owner-file-behavior-differs-between-netdrive-and-synced-fo).
### TL;DR
Nextcloud Sync-Client uses WebDAV for syncing files. Office does not recognise it's own Lock-Files (even if synced). This results in files beeing edited by multiple users at the same time. Last one who saves wins. Collision files are created by Nextcloud, but noone can megre them (merge Visio drawings...yeah right).

In the end I wrote this little wrapper.

# What does this wrapper do?
The wrapper looks for local Microsoft Office Lock-Files (Word, Excel, Visio) of the formats:
* .doc
* .docx
* .docm
* .xlsx (**not .xls**, this is stored somewhere inside Windows Folder)
* .xlsm
* .vsd

If such a file is found, the file is propably opend and must not be opend by another user.

# Requirements to make it work
* Nextcloud Sync Client
* Edit `<Path to Nextcloud-Sync-Client>\sync-exclude.lst`
* Add `*~` to sync-exclude.lst (needs administrative permissions)
* Configure it to sync hidden files