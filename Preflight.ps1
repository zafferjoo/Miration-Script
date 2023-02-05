###Disclaimer: This script does make any change to the environment.
###Ensure Exchange online powershell and msonline modules are installed on the 
###Server with internet connectivity



#This script needs to be customized at following lines.
#Line 57: Server njrarwdc0698.domain.com:3268 needs to be replaced by DCSERVERNAME.DOMAIN.com
#Line 109:$Exchusername ="AD Login"
#Line 110:$pwd="xxxxxx" #AD Password
#Line 111:-ConnectionUri values has to be modified as per your Environment
#Lines 196||380: @domain.mail.onmicrosoft.com must be replaced as per your Tenant 


#Procedure to run this script
#Step1. Create a folder by name 'Migration tasks' on Desktop and place this script there. Name this Script PreFlight.ps1
#Step2. Create a  csv file (Any name say Migrate.csv) with Alias as header and place the Emailaddresses/Alias of the Mailboxes to be scanned.
#Step3. Open AD PowerShell and run the command
#         cd '~\Desktop\Migration tasks'
#Step4: Run the command to start the commmand
#       .\PreFlight.ps1 -Wavefile Migrate.csv -OutPut ExchangeOnprem.csv - Final Report.csv 
# The names of ExchangeOnprem.csv,Report.csv should be changed everytime you run the script
#Step5: Once The Exchange Onprem is scanned it will prompt for credentials. Enter the username of Office 365, Password not required
#Step6: 



[CmdletBinding()]
param(

[Parameter(Mandatory=$True)]
[string]$Wavefile,
 [string]$Output,
 [string]$final


)

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


function User
{
param([string]$Alias)
Write-host "Checking the Status of $Alias"
$User = Get-ADUser -Filter {SamAccountname -eq $alias} -server njrarwdc0698.us.domain.com:3268 -SearchBase "dc=domain,dc=com" -Properties *
if ($user -ne $null)
{
return $user
}
else
{
return $null
}

}

#function Mailbox
function Mailbox
{
param([string]$Alias)
$mail=get-mailbox $alias
if ($mail -ne $null)
{
return $mail
}
else
{
  $mail=get-remotemailbox $alias
  if ($mail -ne $null)
{
return $mail
}
 else
  {
  return $null
  }
}
}



function Stats
{
param([string]$Alias)
$stats=get-mail
return $stats
}




#################SCRIPT STARTS HERE#############################

Import-Module ActiveDirectory
$Exchusername =[string]$a + $env:Username
$pwd="xxxxxx" 
$Exchpassword = $pwd | ConvertTo-SecureString -AsPlainText -Force
$credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "$Exchusername,$Exchpassword"
$Exchsession = New-PSSession -ConfigurationName Microsoft.exchange -ConnectionUri "http://server.domaincom/powershell/" -Authentication Kerberos  -Credential $credentials
Import-PSSession -Session $Exchsession -AllowClobber

if( (get-msoldomain).count -lt 1)
{
connect-msolservice
}

$dom=$null
$dom=@()
$domains =get-msoldomain
foreach($d in $domains)
{
$dom+=$d.name
}


$total=@()
set-adserversettings -viewentireforest 1
$Rec=import-csv $Wavefile

$nosync=$null ;$nosync=@()

foreach($Mail in $rec)
{

$Mdata=Mailbox -Alias $Mail.Alias
$uData= User -Alias $mail.Alias
#$stats= Stats -Alias $mail

if ($Mdata -eq $null -and $Udata -eq $null)
{
$obj=new-object psobject
$obj |add-member -NotePropertyname Exists -NotePropertyValue "NO"
$obj |add-member -NotePropertyname Displayname -NotePropertyValue "NA"
$obj |add-member -NotePropertyname Alias -NotePropertyValue $($Mail.Alias)  
$obj |add-member -NotePropertyname MailboxStatus -NotePropertyValue "NA"
$obj |add-member -NotePropertyname Primarysmtpaddress -NotePropertyValue "NA"
$obj |add-member -NotePropertyname Domainname -NotePropertyValue "NA"
$obj |add-member -NotePropertyname Mailboxtype -NotePropertyValue "NA"
$obj |add-member -NotePropertyname UMenabled -NotePropertyValue "NA"
$obj |add-member -NotePropertyname Country -NotePropertyValue "NA"
$obj |add-member -NotePropertyname PDL -NotePropertyValue "NA"
$obj |add-member -NotePropertyname 	domainLogin -NotePropertyValue "NA"
$obj |add-member -NotePropertyname 	Stamping -NotePropertyValue "NA"
$obj |add-member -NotePropertyname MailboxDataBase -NotePropertyValue "NA"
$obj |add-member -NotePropertyname 	OnPremguid -NotePropertyValue "NA"
$obj |add-member -NotePropertyname CustomAttribute9 -NotePropertyValue "NA"
$obj |add-member -NotePropertyname 	CustomAttribute3 -NotePropertyValue "NA"
#$obj |add-member -NotePropertyname 	CustomAttribute7 -NotePropertyValue "NA"
$obj |add-member -NotePropertyname 	Forwarding -NotePropertyValue "NA"
$obj |add-member -NotePropertyname 	ForwardingAddress -NotePropertyValue "NA"
$obj |add-member -NotePropertyname 	SIPaddress -NotePropertyValue "NA"
$obj |add-member -NotePropertyname 	NAdomain -NotePropertyValue "NA"
$obj |Export-csv -nti $Output -append -force
}

else
{
$obj=new-object psobject
$obj |add-member -NotePropertyname Exists -NotePropertyValue "YES"
$obj |add-member -NotePropertyname Displayname -NotePropertyValue $($Mdata.Displayname)
$obj |add-member -Notepropertyname Alias -NotePropertyValue $Mail.Alias

if ($Mdata.recipienttypedetails  -match "remote")
{
$obj |add-member -NotePropertyname MailboxStatus -NotePropertyValue "MIGRATED"
}

if ($Mdata.recipienttypedetails  -eq "UserMailbox")
{
$obj |add-member -NotePropertyname MailboxStatus -NotePropertyValue "PENDING"
}


$obj |add-member -NotePropertyname Primarysmtpaddress -NotePropertyValue $($Mdata.Primarysmtpaddress)
$obj |add-member -NotePropertyname Domainname -NotePropertyValue $($Mdata.CustomAttribute7)
$obj |add-member -NotePropertyname Mailboxtype -NotePropertyValue $($Mdata.Recipienttypedetails)
$obj |add-member -NotePropertyname UMenabled -NotePropertyValue $($Mdata.umenabled)
$obj |add-member -NotePropertyname Country -NotePropertyValue $($Udata.Country)
$obj |add-member -NotePropertyname PDL -NotePropertyValue $($Udata.'msDS-preferredDataLocation')
$obj |add-member -NotePropertyname domainLogin -NotePropertyValue $($Udata.Userprincipalname)



$alias =$Mdata.alias +"@domain.mail.onmicrosoft.com"
if(($Mdata |?{$_.emailaddresses -match $alias}) -eq $null)
{
$obj |add-member -NotePropertyname Stamping -NotePropertyValue "NO"
}
else
{
$obj |add-member -NotePropertyname Stamping -NotePropertyValue "YES"
}




$obj |add-member -NotePropertyname MailboxDataBase -NotePropertyValue $($Mdata.DataBase)
$obj |add-member -NotePropertyname OnPremguid -NotePropertyValue $($Mdata.Exchangeguid)
$obj |add-member -NotePropertyname CustomAttribute9 -NotePropertyValue $Mdata.CustomAttribute9
$obj |add-member -NotePropertyname CustomAttribute3 -NotePropertyValue $Mdata.CustomAttribute3



$forward=$Mdata.ForwardingSmtpAddress
if(!($forward -eq $null)){

$obj |add-member -NotePropertyname Forwarding -NotePropertyValue "YES"
$obj |add-member -NotePropertyname ForwardingAddress -NotePropertyValue $forward.SmtpAddress


}

else
{
  $obj |add-member -NotePropertyname Forwarding -NotePropertyValue "NO"
  $obj |add-member -NotePropertyname  ForwardingAddress -NotePropertyValue "NULL"
 
}

$sip=($Mdata|select -ExpandProperty emailaddresses |select @{n="Email" ;e={$_}} | ?{$_.email -match "Sip:"}).Email
$obj |add-member -NotePropertyname SIPaddress -NotePropertyValue $sip

$email=$Mdata |select -expandproperty Emailaddresses |select @{n="Email" ; e={$_}}
$badalias=$null
$badalias=@()
foreach($e in $email)
{
$bs=$null
$bs= ($e.email -split "@" )[1]
if($dom -contains $bs)
{
 
}
else
{
if( $e.email -match "x400" -or $e.EMAIL -match "x500")
{

}
else
{
$badalias+=$bs
}
}
}
if ($badalias.count)
{
$obj |add-member -NotePropertyname NAdomain -NotePropertyValue "YES"
Write-host "The user $($Mail.Alias) has an non accepted domain named $($badalias) "
}
else
{
$obj |add-member -NotePropertyname NAdomain -NotePropertyValue "NO"
}
$obj |Export-csv -nti $Output -append -force
}


}


################################################# SCRIPT FOR EXCHANGE ONLINE #################################

Remove-pssession $Exchsession
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

$user=import-csv $Output

Foreach($us in $user)
{

$Op=$us.Alias
Write-Host "Checking th Online Status of user ----> $op" -ForegroundColor Green

$status=Checkconnected

if($status -match "Broken" )
          {
               Write-host "The Connection has broken ......... ____Rejoining again" -ForegroundColor Yellow
               ConnectEXO($cred.Username)
             
           }

################Actual work################
$rec=$us |%{get-recipient $_.Alias}

$Add= new-object psobject
$add =$us
$OnLineStatus=$null
$RemoteMailsync=$null
$Login=$null
$Lic=$null

if($Rec -eq $null)
  {
   ConnectEXO($cred.Username)
   $rec=$us |%{get-recipient $_.Alias}

    if($rec -eq $null)
      {
          $guidMatch= "No SYNC"
          $OnLineStatus="NO SYNC"
          $RemoteMailsync="NO SYNC"
          $Login=" NO SYNC"
          $Lic="NO SYNC"
          $CustomAttribute7="NA"
        }  
  }
else
{



$Login=$rec.windowsLiveid
if($add.onpremguid -eq $rec.exchangeguid)
{
$GuidMatch="YES"
}
else
{
$GuidMatch="NO"
}

$lic=""
$Lic=(Get-MsolUser -UserPrincipalName $Login |select -ExpandProperty licenses |select accountskuid |?{$_.accountskuid -match "domain:ENTERPRISEPACK" -or $_.accountskuid -match "domain:ENTERPRISEPREMIUM"}).ACCOUNTSKUID
if ($lic -match "ENTERPRISEPACK" -or $lic -match "ENTERPRISEPREMIUM")
 { 
    if ($lic -match "ENTERPRISEPREMIUM")         
      {
          $Lic="E5"
       }
   else
      {
       $lic="E3"
        }
  }
   else
{
      $Lic= ""
}



$status=$rec.RecipientType
if($status -eq "Mailuser")
{
$OnLineStatus="Pending"
}
if($status -eq "UserMailbox")
{
$OnLineStatus="Migrated"
}

$REmail=$rec.Alias +"@domain.Mail.onmicrosoft.com"
$email=$rec |select -expandproperty Emailaddresses |select @{n="Email" ; e={$_}} |?{$_.Email -match $remail}
if($email -ne $null)
{
$RemoteMailsync="YES"
}
else
{
$RemoteMailsync="NO"
}
}
$add |add-member -NotePropertyname  GuidMatch -NotePropertyValue $guidMatch
$add |add-member -NotePropertyname  OnlineGuid -NotePropertyValue $rec.exchangeguid
$add |add-member -NotePropertyname  OnLineStatus -NotePropertyValue $OnLineStatus
$add |add-member -NotePropertyname  RemoteMailsync -NotePropertyValue $RemoteMailsync
$add |add-member -NotePropertyname  Login -NotePropertyValue $Login
$add |add-member -NotePropertyname  Licenses -NotePropertyValue $Lic
$add |add-member -NotePropertyname CustomAttribute7 -NotePropertyValue $rec.CustomAttribute7





$add |export-csv -nti $final -append -force
}

Write-Host "Analysis in Progress ........" -foregroundcolor Yellow








$Stat=$null
$stat=@()
$guid=$null
$guid=@()
$RevE5=$null
$RevE5=@()
$UpE3=$null
$UpE3=@()
$Ret=$null
$Ret=@()
$fin=import-csv $final
foreach ($fi in $fin)
{
  if ($fi.MailboxStatus -ne $fi.OnLineStatus)
      {$stat+=$fi }
  if($fi.GuidMatch -eq "No")
      { $guid+=$fi}
  if($fi.Licenses -eq "E5" -and $fi.Umenabled -eq $false)
      { $RevE5+=$fi }
  if($fi.Licenses -eq "E3" -and $fi.Umenabled -eq $true)
      {$UpE3+=$fi}
  if($fi.Customattribute9 -match ":")
     {$Ret+=$fi}
}

$user=import-csv $final
$user |?{$_.Exists -match "no"} |Select Alias, @{n="Error";e={"Mailboxnot found"}} |export-csv Error.csv -append
$user|?{$_.Mailboxtype -match "remote"} |Select Alias, @{n="Error";e={"Mailboxnot already Migrated"}} |export-csv Error.csv -append
$user|?{$_.UMenabled -eq $true}|Select Alias, @{n="Error";e={"UMEnabled"}} |export-csv Error.csv -append
$user|?{$_.CustomAttribute9 -match ":"}|Select Alias, @{n="Error";e={"Mailbox to terminate"}} |export-csv Error.csv -append
$user|?{$_.Forwarding -match "YES"}|Select Alias, @{n="Error";e={"Forwarding Enabled"}} |export-csv Error.csv -append 
$user|?{$_.NAdomain -match "yes" }|Select Alias, @{n="Error";e={"Non accepted Domain"}} |export-csv Error.csv -append
$user|?{$_.GuidMatch   -match  "NO"}|Select Alias, @{n="Error";e={"Mailbox Guid mismatch"}} |export-csv Error.csv -append
$user|?{$_.OnLineStatus  -match "NO SYNC"}|Select Alias, @{n="Error";e={"Mailbox not Syncing"}} |export-csv Error.csv -append
$user|?{$_.CustomAttribute7 -match "retired|terminate"}|Select Alias, @{n="Error";e={"User  to retire|terminated"}} |export-csv Error.csv -append

del Error.csv
Write-host "The Unfit users for migration are present in Error.csv"



