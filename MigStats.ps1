#Place this script in the Migration Scripts folder on the desktop of your machine by name Report.ps1.
#Connect to Exchange online PowerShell and Office365 Admin Powershell 
#Once connected to the office 365 , use the command get-msolacountsku to get the name of the licenses and modify the following lines
#Line 52,54,55 


#Steps to use the script.
#Use the command to get the list of migration batch and save it on notepad.
#        get-migrationbatch
#When scripts starts it prompts for username and Password. Kindly enter office 365 admin creds.
# .\Report.ps1 -Batchfile List.csv -Filename Report.csv
#  The Report.csv gives the list of Migration Report



[CmdletBinding()]
param(

[Parameter(Mandatory=$True)]
[string]$BatchFile,
 [string]$fileName


)

function Itemcount
{
param([string]$user)
$skip=Get-MigrationUserStatistics $user |?{!($_.Status -match "Failed")}
return $skip
}

function ConnectEXO
{
param([string]$user)
get-pssession |remove-pssession
connect-exopssession -userprincipalname $user 
}

function Checkconnected
{
$session=Get-PSSession |?{$_.computername -match "outlook.office365.com"}
$status=$session.state
return $status
}

function Migstats 
{
param([string]$user)
$Mig=Get-MoveRequestStatistics $user
if($mig -eq $null){Write-host "Command Failing for user $user." -ForegroundColor Red}
return $Mig
}  

function Licenses 
{
param([string]$user)
$use=get-msoluser -userprincipalname $user
$sku=($use |select -expandproperty licenses).accountskuid
$sku=$sku |?{$_ -match "domain:ENTERPRISEPREMIUM" -or $_ -match "domain:ENTERPRISEPACK"}
switch ($sku) {
"domain:ENTERPRISEPREMIUM" {$lics="E5" ; return $lics}
"domain:ENTERPRISEPACK" {$lics="E3" ; return $lics}
 default{$lics="No Licenses"; return $lics}
}
}

############################Main Function Starts Here########################

$cred=get-credential
$path= '~\desktop\Migration tasks'
$status=Checkconnected
if($status -eq $null)
{
connect-msolservice
connect-exopssession -userprincipalname $cred.Username
}

if($status -match "Broken" )
          {
               Write-host "The Connection has broken ......... ____Rejoining again" -ForegroundColor Yellow
               ConnectEXO($cred.Username)
             
           }



$Batch=Read-Host "Name of the Migration batch"
$users=get-migrationuser -BatchId $batch -Resultsize unlimited 
$users |select Identity |export-csv $BatchFile
$user=import-csv $BatchFile
$file=$fileName


Foreach($us in $user)
{
$Op=$us.Identity
Write-Host "Checking the Migration Status of user ----> $op"-ForegroundColor Green

$status=Checkconnected

if($status -match "Broken" )
          {
               Write-host "The Connection has broken ......... ____Rejoining again" -ForegroundColor Yellow
               ConnectEXO($cred.Username)
             
           }
$rec=$us |%{get-recipient $_.identity}
if($Rec -eq $null)
  {
   ConnectEXO($cred.Username)
   $rec=$us |%{get-recipient $_.identity}
  }

$x=$rec.alias
$user=$x
 

$MigData=Migstats -user $op
if($MigData -eq $null)
  {
   ConnectEXO($cred.Username)
   $MigData=Migstats -user $op
  }



$skip= Itemcount -user $op
if($skip -eq $null)
{
   ConnectEXO($cred.Username)
   $skip= Itemcount -user $Op
}
$sk=$skip.SkippedItemCount


$office365Upn=($Rec).Windowsliveid
$LicensesInfo= Licenses -User  $office365Upn


$ComItem=$Migdata.PercentComplete


$obj=new-object psobject
$obj |add-member -NotePropertyname Displayname -NotePropertyValue $Rec.Displayname
$obj |add-member -NotePropertyname Primarysmtpaddress -NotePropertyValue $Rec.PrimarySmtpaddress
$obj |add-member -NotePropertyname Alias -NotePropertyValue $Rec.alias
$obj |add-member -NotePropertyname Userprincipalname -NotePropertyValue $office365Upn
$obj |add-member -NotePropertyname Statusdetail -NotePropertyValue $Migdata.Statusdetail
$obj |add-member -NotePropertyname Batchname -NotePropertyValue $us.BatchId
$obj |add-member -NotePropertyname Percentcomplete -NotePropertyValue $ComItem
$obj |add-member -NotePropertyname Licensesname -NotePropertyValue $LicensesInfo
$obj |add-member -NotePropertyname SkipItemcount  -NotePropertyValue $Sk
#$path='~\desktop\Reports\' + $filename
#$filename
$obj |export-csv  -nti $filename  -append
}

Write-Host "The Migration is saved in file $($fileName)" 