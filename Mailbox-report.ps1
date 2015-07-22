# Script will get organization users and print it's name, email address, account usage and last logon time

$ErrorActionPreference = "SilentlyContinue";
$scriptpath = $MyInvocation.MyCommand.Definition 
$dir = Split-Path $scriptpath 

#Variables to configure
$organization = "";

#No change needed from here!!!
$reportPath = "$dir\"
$reportName = "Report_$(get-date -format dd-MM-yyyy__HH-mm-ss).html";
$mbxReport = $reportPath + $reportName

$i = 0;
If (Test-Path $mbxReport)
    {
        Remove-Item $mbxReport
    }

#Loading Exchange CMDlets. Please check Paths in case this doesn't work
. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1' 
Connect-ExchangeServer -auto
cls

Write-Host "Getting Mailbox info..."
$mbxArray = Get-Mailbox -organization $organization -ResultSize Unlimited

$Total = $mbxArray.count
Write-Host "$Total mailboxes will be proccesed"

$titleDate = get-date -uformat "%d-%m-%Y"
$header = "
		<html>
		<head>
		<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
		<title>User report - $organization</title>
		<STYLE TYPE='text/css'>
		<!--
        	table {
            		border: thin solid #666666;
        	}
		td {
			font-family: Tahoma;
			font-size: 11px;
			border-top: 1px solid #999999;
			border-right: 1px solid #999999;
			border-bottom: 1px solid #999999;
			border-left: 1px solid #999999;
			padding-top: 0px;
			padding-right: 0px;
			padding-bottom: 0px;
			padding-left: 0px;
		}
		body {
			margin-left: 5px;
			margin-top: 5px;
			margin-right: 0px;
			margin-bottom: 10px;
			table {
			border: thin solid #000000;
		}
		-->
		</style>
		</head>
		<body>
		<table width='100%'>
		<tr bgcolor='#CCCCCC'>
		<td colspan='7' height='25' align='center'>
		<font face='tahoma' color='#003399' size='4'><strong>User report - $Organization - $titledate</strong></font>
		</td>
		</tr>
		</table>
"
 Add-Content $mbxReport $header
 $tableHeader = "
 <table width='100%'><tbody>
	<tr bgcolor=#CCCCCC>
    <td width='10%' align='center'>User</td>
	<td width='5%' align='center'>Mail</td>
	<td width='15%' align='center'>Mailbox Plan</td>
	<td width='10%' align='center'>Size (MB)</td>
    <td width='10%' align='center'>Last Logon</td>
	</tr>
"
Add-Content $mbxReport $tableHeader
  foreach($mbx in $mbxArray)
	{	
	$Name = $mbx.Name
    $Plan = $mbx.MailboxPlan
    $Address = $mbx.PrimarySMTPAddress
    $Statistics = Get-MailboxStatistics -Identity $mbx.Identity
     
		foreach($stats in $Statistics)
	{        
		$Size = $stats.TotalItemSize.value.ToMB()
        $Logon = $stats.LastLogonTime
        
	$dataRow = "
		<tr>
        <td width='10%' align='center'>$Name</td>
		<td width='5%'>$Address</td>
		<td width='15%' align='center'>$Plan</td>
		<td width='10%' align='center'>$Size</td>
		<td width='10%' align='center'>$Logon</td>
		</tr>
"
Add-Content $mbxReport $dataRow;
$i++
Write-Host "Proccesing $Address ($i of $Total)";
	}
}

Add-Content $mbxReport  "</body></html>"
