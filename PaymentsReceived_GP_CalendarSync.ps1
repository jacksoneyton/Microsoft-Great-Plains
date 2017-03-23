function Invoke-SQL {
    param(
        [string] $dataSource = ".\SQLEXPRESS",
        [string] $database = "MasterData",
        [string] $sqlCommand = $(throw "Please specify a query.")
      )

    $connectionString = "Data Source=$dataSource; " +
            "Integrated Security=SSPI; " +
            "Initial Catalog=$database"

    $connection = new-object system.data.SqlClient.SQLConnection($connectionString)
    $command = new-object system.data.sqlclient.sqlcommand($sqlCommand,$connection)
    $connection.Open()

    $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $command
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataSet) | Out-Null

    $connection.Close()
    $dataSet.Tables

}

$SQLQuery = @"
SELECT dbo.RM20101.CUSTNMBR as CustomerNumber, dbo.Customers.[Customer Name] as CustomerName , dbo.RM20101.ORTRXAMT as TransactionAmt, dbo.RM20101.DOCDATE as DatePaid
FROM dbo.RM20101
JOIN dbo.Customers on (dbo.customers.[Customer Number] = dbo.RM20101.CUSTNMBR)
WHERE (dbo.rm20101.RMDTYPAL = 9) AND (dbo.rm20101.VOIDSTTS = 0) AND (dbo.rm20101.DOCDATE > DATEADD(day, -1, GETDATE()))
"@



## Load Managed API dll  

###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
if (Test-Path $EWSDLL){Import-Module $EWSDLL} else{
    echo "$(get-date -format yyyyMMddHHmmss):"
    echo "This script requires the EWS Managed API 1.2 or later."
    echo "Please download and install the current version of the EWS Managed API from"
    echo "http://go.microsoft.com/fwlink/?LinkId=255472"
    echo ""
    echo "Exiting Script."
    Start-Sleep -Seconds 3
    exit
    } 

# SET EWS URI -- EDIT THIS!!!
$ExchangeWebServicesURL = "https://<yourmailserver.com>/ews/Exchange.asmx"

# Create a new Exchange Service Object
$exchService = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010)

# Set the Credentials -- EDIT THIS!!!
$exchService.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials('youruser','yourpass')

#Set the URL for the service
$exchService.Url= new-object Uri($ExchangeWebServicesURL)

# Bind to the Calendar folder  
# $folderid generates a true
$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)   
$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService,$folderid)

# PULL DATA FROM GP DATABASE -- EDIT THIS!!!
$SQLResult = Invoke-SQL -dataSource "yourGPserver" -database "Your_GP_DB" -sqlCommand $SQLQuery

# CREATE APPOINTMENTS
foreach ($result in $SQLResult)
{
    $appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment -ArgumentList $exchService
        $appointment.Subject = "$($result.CustomerName)" + " - " + "$($result.CustomerNumber)"
        $appointment.Body = "$($result.CustomerName)" + " Paid " + "$($result.TransactionAmt)" + " on " + "$($result.DatePaid)"
        $appointment.Start = "$($result.DatePaid)"
        $appointment.End = "$($result.DatePaid)"
    $appointment.Save()
}
