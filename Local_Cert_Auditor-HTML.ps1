#region Modules Reduired

Import-Module WebAdministration -ErrorAction SilentlyContinue
Import-Module ActiveDirectory -force -ErrorAction SilentlyContinue

#endregion modules required#endregion modules required

#region Adjustable Variables - Please adjust these variables to your enviorment to get the information you need from the report.


#Change the $DaystoSearch variable for how many days you would like to know in advance for the expiring certificates.
$DaystoSearch = "<Enter amount of days you would like to search for the certificate to expire in the future>"

#Change the $Servers variable to a varaint of Get-ADComputers to ensure you get a list of servers you are inquiring upon. Can be more or less restrictive.
$servers = Get-ADComputer -Filter "OperatingSystem -like 'Windows Server*' -and Enabled -eq '$true'" | Sort-Object |Foreach {$_.Name}

#Change the $Path variable to the path where you would like Doc Report to be saved.
$Path = "<Enter a path to a centrally saved location for the report>"

#Change the $DOCReport variable to the naming convention of your choosing - As set below this will create a new folder for each Month/Year for historical purposes ie \\fileserver\2022-June\180 Day Report.html
$DOCreport = "$Path\$YearMonth\$currentDay - $DaystoSearch Days Certificate Expiration Report.htm"

#endregion Adjustable Variables

#region Static Variables - These variables are not required to be adjusted as they are just formatting for later in the script. Can adjust the Get-Date format if not liking MM/dd/yyyy format

#Sets variables for the enviorment
$today = ((Get-Date).AddDays($DaystoSearch)).tostring("MM-dd-yyyy")
$otoday = Get-Date -Format "MM/dd/yyy"
$currentDay = Get-Date -Format "MM-dd-yyyy"
$currentMonth = Get-Date -UFormat %m
$currentYear = Get-Date -Format "yyyy"
$currentMonth = (Get-Culture).DateTimeFormat.GetMonthName($currentMonth)
$MonthYear = $currentMonth + '-' + $currentYear
$YearMonth = $currentYear + '-' + $currentMonth
$fullpath = "$Path\$YearMonth"

#endregion Static Variables

#region Create New Report

#Checks if the path for the year month exists, if not creates it.
#Uncomment the write-host lines if running manually to keep updated on the process.

if (Test-Path "$Path\$YearMonth")
    {
    #Write-Host "Path Exists"
    }

else
    {
    #Write-Host "Path does not exist. Creating path now."
    New-Item -Path $Path -Name $YearMonth -ItemType Directory
    }

#Creates a new report for each month/year combination
New-Item -Path $fullpath -Name "$currentDay - $DaystoSearch Days Certificate Expiration Report.htm" -ItemType File

#endregion of new report
 
#region HTML Styling with CSS
#CSS Style config

$a = @"
<style>

    h1 {

        font-family: Arial, Helvetica, sans-serif;
        color: #e68a00;
        font-size: 28px;

    }

    
    h2 {

        font-family: Arial, Helvetica, sans-serif;
        color: #000099;
        font-size: 16px;

    }

    
    
   table {
		font-size: 12px;
		border: 0px; 
		font-family: Arial, Helvetica, sans-serif;
	} 
	
    td {
		padding: 4px;
		margin: 0px;
		border: 0;
	}
	
    th {
        background: #395870;
        background: linear-gradient(#49708f, #293f50);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
	}

    tbody tr:nth-child(even) {
        background: #f0f0f2;
    }
    


    #CreationDate {

        font-family: Arial, Helvetica, sans-serif;
        color: #ff3300;
        font-size: 12px;

    }





</style>

<title>Certificate Auditor</title>

"@
#endregion

#region Add Info to top of HTML Report 

#Adds headers and other content to HTML report

#$i=0
ConvertTo-Html -head $a >> $DOCreport
ConvertTo-Html -PreContent "<h1>Certificate Auditor - $DaystoSearch Day Audit</h1><h2>Please work to resolve any expiring certs in this report if necessary.</h2>" >> $DOCreport

#endregion

#region Create table to input to report

#Create the array for the info to be submitted to
$expired = @()

#region Pull Cert Information from all $servers

#Run through each computer and check for certs within the configured date period
foreach ($server in $servers) { #each server start
    Write-Host "Working on:" $server
	$certs = Invoke-Command -ComputerName $server -ScriptBlock { Get-ChildItem -path Cert:\LocalMachine\My -Recurse -erroraction SilentlyContinue | select-Object Subject, Issuer, FriendlyName, Thumbprint, NotBefore, NotAfter  -ExcludeProperty PSComputerName, RunspaceId, PSSHowComputerName | Sort-Object Notafter -Descending } -ErrorAction SilentlyContinue
    
    foreach ($c in $certs) { #each cert start
        $subs = $c.subject.split(",")
        foreach ($sub in $subs) {
            if ($sub -like "*CN=*") {
                $sj = $sub -replace 'CN=', ''
                $SJ = $sj -replace '\s',''
                }
        }
        $edate = $c.NotAfter.ToString("MM-dd-yyyy")
        $ts = New-TimeSpan -Start $edate -End $today
        if ($ts.days -gt 0 -and $ts.Days -lt $DaystoSearch) {
            $item = [PSCustomobject]@{
                Name = $server
                Issuer = $c.Issuer
                Thumbprint = $c.Thumbprint
                'Issued Date' = $c.NotBefore.ToString("MM-dd-yyyy")
                'Expiration Date' = $c.NotAfter.tostring("MM-dd-yyyy")
            }
            $expired = $expired + $item
        }
    } #each cert end
} #each server end

#endregion

#Format the array to a table
$sortexpired = $expired | Sort-Object Name
$sortexpired | Format-Table -AutoSize

#endregion

#region Update HTML Report

#Takes the table created from array and puts it into the table for the HTML
$sortexpired | ConvertTo-Html -head $a -property Name,Issuer,Thumbprint,'Issued Date','Expiration Date' >> $DOCreport

#Adds footer information to the report
ConvertTo-Html -PostContent "<p>Report Creation Date: $(Get-Date)</p>" >> $DOCreport

#endregion

#region Email report

#Sends an email over your SMTP server to a person or distribution list

$From = "CertificateAuditor@DOMAIN.com"
$To = "<Enter the address you want to send an email to when report is finished>"
$Attachment = "$DOCreport"
$Subject = "Certificates expiring within $DaystoSearch days"
$Body = "<h3><u>Certificates expiring within the next $DaystoSearch days.</u></h3><p>Please find attached the HTML report detailing certificates that will expire within the next $DaystoSearch days.<BR>Please review the report and make relevant arrangements to replace, or renew these certificates to reduce the risk of a service outage.  A copy of this report has alerady been saved centrally at '$DOCreport'</p></br/>Brought to you by Certificate Auditor Script"
$SMTPServer = "<Enter FQDN of SMTP server here>"
$SMTPPort = "25"
Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -port $SMTPPort -Attachments $Attachment –DeliveryNotificationOption OnSuccess

#endregion Email Report
