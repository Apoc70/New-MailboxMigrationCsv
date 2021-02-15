<#
    .SYNOPSIS
    Export primary email addresses of Exchange mailboxes as Csv in preparation for mailbox migration batches.
   
    Thomas Stensitzki
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.1, 2021-02-15

    Ideas, comments, and suggestions via GitHub.
 
    .LINK  
    http://www.granikos.eu/en/scripts
   
	
    .DESCRIPTION
	
    This script exports the primary email address of Exchange mailboxes as a single Csv file using "EmailAddress" as column header. 
    The scripts splits the Csv-file into batch-size Csv files. You can use these Csv files as input files when creating mailbox migration batches.
    You can define number of email address per splitted Csv file.
    The scripts allows for exporting a selected mailbox recipient type or all mailbox types.

    Tested with Exchange Server 2016 and Exchange Server 2019.

    .NOTES 
    Requirements 
    - Windows Server 2016+     
    - Administrative Exchange Server 2016+ Management Shell

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0 Initial community release 
    
	
    .PARAMETER MailboxType
    Switch to select the mailbox type you want to export.
    Options are User, Shared, Room, Arbitration, PublicFolder, Equipment, All

    .PARAMETER BatchSize
    The number of email addresses per split CSV file, aka batch size.
    Default: 25

    .PARAMETER ViewEntireForest
    Switch parameter to set the search scopt to the entire Active Directory forest.
    The switch might be required when exporting arbitration and system mailboxes and the AD forest uses multi-domain infrastructure. 

    .PARAMETER UseBatchFolder
    Switch parameter to store splitted Csv files in a sub-folder named \Batches

    .EXAMPLE
    Export all user mailboxes and create Csv files of 25 users per file, store Csv files in the script folder
    .\New-MailboxMigrationCsv.ps1 -MailboxType User

    .EXAMPLE
    Export all room mailboxes and create Csv files of 75 mailboxes per file, store Csv files in \Batches folder.
    .\New-MailboxMigrationCsv.ps1 -MailboxType Room -BatchSize 75 -UseBatchFolder

#>

param (
    [Parameter(Mandatory,HelpMessage='Please select a mailbox type or select All')]
    [ValidateSet('User','Shared','Room','Arbitration','PublicFolder','Equipment','All')]
    [string]$MailboxType,
    [int]$BatchSize = 25,
    [switch]$ViewEntireForest,
    [switch]$UseBatchFolder
  )
  
  function Split-CsvFile {
    <#
        .SYNOPSIS
        This function splits a single Csv file into multiple fiels preserving the original Csv header
  
        .DESCRIPTION
        We need to split the exported Csv file with all mailbox email addresses into separate Csv files. We use these files for creating the mailbox migration batches.
  
        .PARAMETER MasterFile
        The Csv sourve file that you want to split
  
        .PARAMETER BatchFilename
        The filename pattern for the resulting Csv files
  
        .PARAMETER BatchRowSize
        Describe parameter -BatchRowSize.
  
        .EXAMPLE
        Split-CsvFile -MasterFile C:\Script\ExportedMailboxes,csv -BatchFilename UserBatch -BatchRowSize 25
        The function splits the Csv file C:\Script\ExportedMailboxes,csv into multiple Csv files containing 25 email adresses. The file name will be UserBatch1.csv etc.
  
        .NOTES
        Thanks to Dustin Dortch https://dustindortch.com/
  
        .LINK
        http://scripts.granikos.eu
    #>
  
  
    param(
      [Parameter(Mandatory,HelpMessage='Provide a CSV source file')]
      [string]$MasterFile,
      [Parameter(Mandatory,HelpMessage='Provide a file name for the split CSV files')]
      [string]$BatchFilename,
      [int]$BatchRowSize
    )
  
    Write-Host ('Splitting CSV masterfile {0}, batch size {1}' -f $MasterFile, $BatchRowSize)
  
    # Default folder name for configured CSV batch files
    $BatchFolderName = 'Batches'
  
    if($UseBatchFolder) {
      if(-not (Test-Path -Path (Join-Path -Path $scriptPath -ChildPath $BatchFolderName))) {
        # Create batch folder 
        $null = New-Item -Path (Join-Path -Path $scriptPath -ChildPath $BatchFolderName) -ItemType Directory -Force -Confirm:$false
      }
      $BatchFolderPath = Join-Path -Path $scriptPath -ChildPath $BatchFolderName
    }
    else {
      $BatchFolderPath = $scriptPath
    }
  
    # Import CSV file
    $CSVFile = Import-Csv -Path $MasterFile -Encoding UTF8
  
    # Calculate the number of batch files
    if($CSVFile.Count -lt $BatchRowSize) {
      $BatchRowSize = $CSVFile.Count + 1
      $Count = 1
    }
    else {
      $Count = [math]::ceiling($CSVFile.Count/$BatchRowSize)
    }
      
    $Length = 'D' + ([math]::floor([math]::Log10($Count) + 1)).ToString()
      
    1 .. $Count | ForEach-Object {
  
      $i = $_.ToString($Length)
      $Offset = $BatchRowSize * ($_ - 1)
  
      $CsvBatchFile = Join-Path -Path $BatchFolderPath -ChildPath ('{0}{1}.csv' -f $BatchFilename, $i)
  
      If(Test-Path -Path $CsvBatchFile) {
        Write-Warning  -Message ('Deleting existing CSV batch file {0}!' -f $CsvBatchFile)
        Remove-Item -Path $CsvBatchFile -Confirm:$false -Force
      }
          
      # Export to file
      If($_ -eq 1) {
        $CSVFile | Select-Object -First $BatchRowSize | Export-Csv -Path $CsvBatchFile -NoTypeInformation -NoClobber 
      } Else {
        $CSVFile | Select-Object -First $BatchRowSize -Skip $Offset | Export-Csv -Path $CsvBatchFile -NoTypeInformation -NoClobber
      }
    }
  }
  
  function Get-Mailboxes {
    <#
        .SYNOPSIS
        Get all mailboxes of given RecipientTypeDetails property and export as Csv file
  
        .DESCRIPTION
        This function fetches all mailboxes of a fiven RecipientTypeDetails property and exports the PrimarySmtpAddress attribute as EmailAddress into a single CSV file.
  
        .PARAMETER RecipientTypeDetails
        The Exchange RecipientTypeDetails property you want to fetch from Exchange, e.g., UserMailbox
  
        .EXAMPLE
        Get-Mailboxes -RecipientTypeDetails PublicFolderMailbox
        Exports all email address for public folder mailboxes to a single CSV file
  
        .LINK
        http://scripts.granikos.eu
  
    #>
  
  
    [CmdletBinding()]
    param (
      [string]$RecipientTypeDetails
    )
  
    # Export user mailboxes only
    $CsvFileFilepath = Join-Path -Path $scriptPath -ChildPath ('{0}Master.csv' -f $RecipientTypeDetails)
    $BatchOutputFilename = ('{0}Batch' -f $RecipientTypeDetails)
  
    if(Test-Path -Path $CsvFileFilepath) {
      Write-Warning  -Message ('Deleting existing CSV file {0}!' -f $CsvFileFilepath)
      Remove-Item -Path $CsvFileFilepath -Confirm:$false -Force
    }
  
    switch($RecipientTypeDetails) {
      'PublicFolderMailbox' {
        Get-Mailbox -PublicFolder -ResultSize Unlimited | Sort-Object -Property Name | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} | `
        Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force
        # Add some well-know mailboxes
      }
      'ArbitrationMailbox' {
        Get-Mailbox -Arbitration -ResultSize Unlimited | Sort-Object -Property Name | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} | `
        Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force
              
        # Add some well-know mailboxes
        # Not axactly arbitration mailboxes, but still system mailboxes of some sort
        Get-Mailbox -AuditLog | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} | `
        Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force -Append
  
        Get-Mailbox 'Discovery*' | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} | `
        Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force -Append 
      }
      default {
        Get-Mailbox -ResultSize Unlimited | ?{$_.RecipientTypeDetails -eq $RecipientTypeDetails} | Sort-Object -Property Name | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} | `
        Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force
      }
    }
  
    Split-CsvFile -MasterFile $CsvFileFilepath -BatchFilename $BatchOutputFilename -BatchRowSize $BatchSize
  }
  
  #region Main
  
  # Script Path
  $scriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Path
  
  if($ViewEntireForest) {
    Set-ADServerSettings -ViewEntireForest:$true
  }
  
  switch ($MailboxType) {
    'User' {
      Get-Mailboxes -RecipientTypeDetails 'UserMailbox'
    }
    'Shared' {
      Get-Mailboxes -RecipientTypeDetails 'SharedMailbox'
    }
    'Room' { 
      Get-Mailboxes -RecipientTypeDetails 'RoomMailbox'
    }
    'Arbitration' { 
      Get-Mailboxes -RecipientTypeDetails 'ArbitrationMailbox'
    }
    'PublicFolder' { 
      Get-Mailboxes -RecipientTypeDetails 'PublicFolderMailbox'
    }
    'Equipment' {
      Get-Mailboxes -RecipientTypeDetails 'EquipmentMailbox'  
    }
    Default {
      # Export all
      Get-Mailboxes -RecipientTypeDetails 'UserMailbox'
      Get-Mailboxes -RecipientTypeDetails 'SharedMailbox'
      Get-Mailboxes -RecipientTypeDetails 'RoomMailbox'
      Get-Mailboxes -RecipientTypeDetails 'ArbitrationMailbox'
      Get-Mailboxes -RecipientTypeDetails 'PublicFolderMailbox'
      Get-Mailboxes -RecipientTypeDetails 'EquipmentMailbox'
    }
  }
  
  #endregion 