# New-MailboxMigrationCsv

Export primary email addresses of Exchange mailboxes as Csv in preparation for mailbox migration batches.

## Description

This script exports the primary email address of Exchange mailboxes as a single Csv file using "EmailAddress" as column header. 
The scripts splits the Csv-file into batch-size Csv files. You can use these Csv files as input files when creating mailbox migration batches.
You can define number of email address per splitted Csv file.
The scripts allows for exporting a selected mailbox recipient type or all mailbox types.

Tested with Exchange Server 2016 and Exchange Server 2019.

## Requirements

- Windows Server 2016+
- Administrative Exchange Server 2016+ Management Shell

## Parameters

### Parameter MailboxType

Switch to select the mailbox type you want to export.
Options are User, Shared, Room, Arbitration, PublicFolder, Equipment, All

### Database

If set, the script fetches only mailboxes stored in that database.

### BatchSize

The number of email addresses per split CSV file, aka batch size.
Default: 25

### ViewEntireForest

Switch parameter to set the search scopt to the entire Active Directory forest.
The switch might be required when exporting arbitration and system mailboxes and the AD forest uses multi-domain infrastructure.

### UseBatchFolder

Switch parameter to store splitted Csv files in a sub-folder named \Batches

## Examples

``` PowerShell
.\New-MailboxMigrationCsv.ps1 -MailboxType User
```

Export all user mailboxes and create Csv files of 25 users per file, store Csv files in the script folder

``` PowerShell
.\New-MailboxMigrationCsv.ps1 -MailboxType Room -BatchSize 75 -UseBatchFolder
```

Export all room mailboxes and create Csv files of 75 mailboxes per file, store Csv files in \Batches folder.

``` PowerShell
.\New-MailboxMigrationCsv.ps1 -MailboxType Room -BatchSize 20 -UseBatchFolder -Database DB01
```

Export all room mailboxes of database DB01, create Csv files of 20 mailboxes per file, and store Csv files in \Batches folder.

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

## Stay connected

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)
- MVP Blog: [https://blogs.msmvps.com/thomastechtalk/](https://blogs.msmvps.com/thomastechtalk/)
- Tech Talk YouTube Channel (DE): [http://techtalk.granikos.eu](http://techtalk.granikos.eu)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)