<#
    .SYNOPSIS
    Export primary email addresses of Exchange mailboxes as Csv in preparation for mailbox migration batches.
   
    Thomas Stensitzki
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.3, 2021-08-25

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
    1.2 Database parameter added, export handling optimized
    1.3 Archive mailbox support added
    
	
    .PARAMETER MailboxType
    Switch to select the mailbox type you want to export.
    Options are User, Shared, Room, Arbitration, PublicFolder, Equipment, All

    .PARAMETER Database
    If set, the script fetches only mailboxes stored in that database.


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

    .EXAMPLE
    Export all room mailboxes of database DB01, create Csv files of 20 mailboxes per file, and store Csv files in \Batches folder.
    .\New-MailboxMigrationCsv.ps1 -MailboxType Room -BatchSize 20 -UseBatchFolder -Database DB01

#>

param (
  [Parameter(Mandatory,HelpMessage='Please select a mailbox type or select All')]
  [ValidateSet('User','Shared','Room','Arbitration','PublicFolder','Equipment','All','Archive')]
  [string]$MailboxType,
  [string]$Database = '',
  [int]$BatchSize = 25,
  [switch]$ViewEntireForest,
  [switch]$UseBatchFolder
)

$scriptVersion = '1.3'
  
# Script Path
$scriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Path
  
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

  if(Test-Path -Path $MasterFile) { 
  
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
  else {
    Write-Verbose -Message ('Master file {0} not found. Skip splitting.' -f $MasterFile)
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
  
  # Set some variables
  $CsvFileFilepath = Join-Path -Path $scriptPath -ChildPath ('{0}Master.csv' -f $RecipientTypeDetails)
  $BatchOutputFilename = ('{0}Batch' -f $RecipientTypeDetails)

  if($Database -ne '') {    
    $BatchOutputFilename = ('{1}-{0}Batch' -f $RecipientTypeDetails, $Database.ToUpper())
  }
  
  if(Test-Path -Path $CsvFileFilepath) {
    Write-Warning  -Message ('Deleting existing CSV file {0}!' -f $CsvFileFilepath)
    Remove-Item -Path $CsvFileFilepath -Confirm:$false -Force
  }

  if($Database -eq '') {
    # Default exportes across all mailboxes, no dedicated mailbox database specified

    switch($RecipientTypeDetails) {
      'PublicFolderMailbox' {
        $Mailboxes = Get-Mailbox -PublicFolder -ResultSize Unlimited | Sort-Object -Property Name | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}}           

        if (($Mailboxes | Measure-Object).Count -gt 0) {
          Write-Verbose -Message ('Exporting {0} arbitration mailbox(es)' -f ($Mailboxes | Measure-Object).Count)
          $Mailboxes | Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force
        }
        else {
          Write-Verbose -Message ('No arbitration mailbox found in database {0}' -f $Database)
        } 
      }
      'ArbitrationMailbox' {
        $Mailboxes = Get-Mailbox -Arbitration -ResultSize Unlimited | Sort-Object -Property Name | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} 
          
        if (($Mailboxes | Measure-Object).Count -gt 0) {
          Write-Verbose -Message ('Exporting {0} arbitration mailbox(es)' -f ($Mailboxes | Measure-Object).Count)
          $Mailboxes | Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force
        }
        else {
          Write-Verbose -Message ('No arbitration mailbox found in database {0}' -f $Database)
        }  
                
        # Add some well-know mailboxes
        # Not axactly arbitration mailboxes, but still system mailboxes of some sort
        $Mailboxes = Get-Mailbox -AuditLog | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} 
          
        if (($Mailboxes | Measure-Object).Count -gt 0) {
          Write-Verbose -Message 'Exporting AuditLog Mailbox information'
          $Mailboxes | Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force -Append 
        }
        else {
          Write-Verbose -Message ('No AuditLog mailbox found in database {0}' -f $Database)
        }
    
        $Mailboxes = Get-Mailbox 'Discovery*' | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} 
          
        if (($Mailboxes | Measure-Object).Count -gt 0) {
          Write-Verbose -Message 'Exporting Discovery Mailbox information'
          $Mailboxes | Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force -Append 
        }
        else {
          Write-Verbose -Message ('No Discovery mailbox found in database {0}' -f $Database)
        }          
      }
      'UserArchiveMailbox' {
        $Mailboxes = Get-Mailbox -ResultSize Unlimited -Archive | Sort-Object -Property Name | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} 

        if (($Mailboxes | Measure-Object).Count -gt 0) {
          Write-Verbose -Message ('Exporting {0} archive mailbox(es)' -f ($Mailboxes | Measure-Object).Count)
          $Mailboxes | Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force
        }
        else {
          Write-Verbose -Message ('No mailbox found in database {0}' -f $Database)
        }
      }
      default {
        $Mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object{$_.RecipientTypeDetails -eq $RecipientTypeDetails} | Sort-Object -Property Name | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} 
          
        if (($Mailboxes | Measure-Object).Count -gt 0) {
          Write-Verbose -Message ('Exporting {0} mailbox(es)' -f ($Mailboxes | Measure-Object).Count)
          $Mailboxes | Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force
        }
        else {
          Write-Verbose -Message ('No mailbox found in database {0}' -f $Database)
        }
      }
    }  
  }
  else {
    # we need to fetch only mailbox in a dedicated database 

    # check if database exists
    if($null -eq (Get-MailboxDatabase -Identity $Database.Trim() -ErrorAction SilentlyContinue)) {
      Write-Error ('Cannot find database {0}. Please verify database name.' -f $Database)
      exit
    }
    else {
      # just trim the database name
      $Database = $Database.Trim()
    }

    switch($RecipientTypeDetails) {
      'PublicFolderMailbox' {
        $Mailboxes = Get-Mailbox -PublicFolder -ResultSize Unlimited -Database $Database | Sort-Object -Property Name | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} 

        if (($Mailboxes | Measure-Object).Count -gt 0) {
          Write-Verbose -Message ('Exporting {0} public folder mailbox(es)' -f ($Mailboxes | Measure-Object).Count)
          $Mailboxes | Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force
        }
        else {
          Write-Verbose -Message ('No public folder mailbox found in database {0}' -f $Database)
        }  
          
        # Add some well-know mailboxes
      }
      'ArbitrationMailbox' {
        $Mailboxes = Get-Mailbox -Arbitration -ResultSize Unlimited -Database $Database | Sort-Object -Property Name | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} 

        if (($Mailboxes | Measure-Object).Count -gt 0) {
          Write-Verbose -Message ('Exporting {0} arbitration mailbox(es)' -f ($Mailboxes | Measure-Object).Count)
          $Mailboxes | Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force
        }
        else {
          Write-Verbose -Message ('No arbitration mailbox found in database {0}' -f $Database)
        }          
                
        # Add some well-know mailboxes
        # Not axactly arbitration mailboxes, but still system mailboxes of some sort
        $Mailboxes = Get-Mailbox -AuditLog -Database $Database | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} 

        if (($Mailboxes | Measure-Object).Count -gt 0) {
          Write-Verbose -Message 'Exporting AuditLog Mailbox information'
          $Mailboxes | Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force -Append 
        }
        else {
          Write-Verbose -Message ('No AuditLog mailbox found in database {0}' -f $Database)
        }
    
        $Mailboxes = Get-Mailbox 'Discovery*' | Where-Object{$_.Database -eq 'INDB02'} | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} 

        if (($Mailboxes | Measure-Object).Count -gt 0) {
          Write-Verbose -Message 'Exporting Discovery Mailbox information'
          $Mailboxes | Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force -Append 
        }
        else {
          Write-Verbose -Message ('No Discovery mailbox found in database {0}' -f $Database)
        }
      }
      'UserArchiveMailbox' {
		Write-Warning -Message 'Fetching archive mailboxes takes some minutes'
        $Mailboxes = Get-Mailbox -ResultSize Unlimited -Archive | ?{$_.ArchiveDatabase -like $Database} | Sort-Object -Property Name | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} 

        if (($Mailboxes | Measure-Object).Count -gt 0) {
          Write-Verbose -Message ('Exporting {0} archive mailbox(es)' -f ($Mailboxes | Measure-Object).Count)
          $Mailboxes | Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force
        }
        else {
          Write-Verbose -Message ('No mailbox found in database {0}' -f $Database)
        }
      }
      default {
        $Mailboxes = Get-Mailbox -ResultSize Unlimited -Database $Database | Where-Object{$_.RecipientTypeDetails -eq $RecipientTypeDetails} | Sort-Object -Property Name | `
        Select-Object -Property @{Name = 'EmailAddress'; Expression = {$_.PrimarySmtpAddress}} 

        if (($Mailboxes | Measure-Object).Count -gt 0) {
          Write-Verbose -Message ('Exporting {0} mailbox(es)' -f ($Mailboxes | Measure-Object).Count)
          $Mailboxes | Export-Csv -Path $CsvFileFilepath -Delimiter ',' -Encoding UTF8 -NoTypeInformation -NoClobber -Force
        }
        else {
          Write-Verbose -Message ('No mailbox found in database {0}' -f $Database)
        }
      }
    }  
  }
  
  
  Split-CsvFile -MasterFile $CsvFileFilepath -BatchFilename $BatchOutputFilename -BatchRowSize $BatchSize
}
  
#region Main

Write-Verbose -Message ('Starting script - Version {0}' -f $scriptVersion)
  
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
  'Archive' {
    Get-Mailboxes -RecipientTypeDetails 'UserArchiveMailbox'
  }
  Default {
    # Export all besides archives
    Get-Mailboxes -RecipientTypeDetails 'UserMailbox'
    Get-Mailboxes -RecipientTypeDetails 'SharedMailbox'
    Get-Mailboxes -RecipientTypeDetails 'RoomMailbox'
    Get-Mailboxes -RecipientTypeDetails 'ArbitrationMailbox'
    Get-Mailboxes -RecipientTypeDetails 'PublicFolderMailbox'
    Get-Mailboxes -RecipientTypeDetails 'EquipmentMailbox'
  }
}
  
#endregion 