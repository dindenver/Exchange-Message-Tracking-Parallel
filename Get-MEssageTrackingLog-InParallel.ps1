	[CmdletBinding(DefaultParametersetName="FileName")]
param (
        [Parameter(Position = 0, Mandatory = $false, HelpMessage="Enter the Sender`'s SMTP address")]
		[String]
        [Alias("From")]
        [Alias("mail")]
        [Alias("PrimarySMTPAddress")]
        [Alias("WindowsEmailAddress")]
        [Alias("Identity")]
$Sender=$null,
        [Parameter(Position = 1, Mandatory = $false, HelpMessage="Enter the Recipient`'s SMTP address")]
		[String]
        [Alias("To")]
        [Alias("User")]
$Recipient=$null,
        [Parameter(Position = 2, Mandatory = $false, HelpMessage="Enter the earliest event time stamp to search for")]
		[DateTime]
$Start=(Get-Date).AddDays(-31),
        [Parameter(Position = 3, Mandatory = $false, HelpMessage="Enter the earliest event time stamp to search for")]
		[DateTime]
$End=$(Get-Date),
        [Parameter(Position = 4, Mandatory = $false, HelpMessage="Enter the Subject to search for")]
		[String]
        [Alias("Title")]
$Subject=$null,
        [Parameter(ParameterSetName="PSO",Position = 5, Mandatory = $true, HelpMessage="Return a collection of PSObjects")]
		[Switch]
        [Alias("PSObject")]
        [Alias("Collection")]
$PSO,
        [Parameter(ParameterSetName="GUI",Position = 5, Mandatory = $true, HelpMessage="Return a GUI grid view")]
		[Switch]
        [Alias("Grid")]
        [Alias("Window")]
$GUI,
        [Parameter(ParameterSetName="FileName",Position = 5, Mandatory = $false, HelpMessage="Enter the file name to store the results in")]
		[String]
        [Alias("File")]
        [Alias("FilePath")]
        [Alias("Path")]
        [Alias("Name")]
$Filename="$($pwd.path)`\MessageTracking-$(GC ENV:Username)-$(get-date -format MMddyy).CSV"
)

# Initialization
# Does not truncate output
$FormatEnumerationLimit=-1
[String[]]$EXServers=@()

Workflow Main
	{
	param (
		[String[]]
	$EXServers,
		[String]
	$Sender,
		[String]
	$Recipient,
		[DateTime]
	$Start,
		[DateTime]
	$End,
		[String]
	$Subject
	)
# Search message tracking logs on each server.
	foreach -parallel ($EXServer IN $EXServers)
		{
		[array]$exdata += InlineScript
			{
			$isdata=@()
			Function Track-ExMessage
				{
				param (
					[String]
				$EXServer,
					[String]
				$Sender,
					[String]
				$Recipient,
					[DateTime]
				$Start,
					[DateTime]
				$End,
					[String]
				$Subject
				)
				$fndata=@()
write-verbose "Processing $EXServer"
# $Sender has to be $null if it is not specified "" only matches blank sender, $null matches all senders
				if ($Sender -eq "") {Remove-Variable Sender;$Sender = $null}
# $Recipient has to be $null if it is not specified "" only matches blank Recipient, $null matches all Recipients
				if ($Recipient -eq "") {Remove-Variable Recipient;$Recipient = $null}
# $Subject has to be $null if it is not specified "" only matches blank Subject, $null matches all Subjects
				if ($Subject -eq "") {Remove-Variable Subject;$Subject = $null}
				Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
# Loads Remote Management tools
#	if (test-path 'C:\Program Files\Microsoft\Exchange Server\V14\bin\Exchange.ps1') {. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\Exchange.ps1' -erroraction silentlyContinue | out-null}
# 	if (test-path 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1') {. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1' -erroraction silentlyContinue | out-null;Connect-ExchangeServer -AllowClobber -ServerFqdn $EXServer | out-null}
write-verbose "Tracking on $($EXServer)"
				$fndata+=Get-MessageTrackingLog -Server $EXServer -Sender $Sender -Recipient $Recipient -Start $Start -End $End -MessageSubject $Subject -ResultSize Unlimited
write-verbose "Tracking on $($EXServer) successful - $?"
write-verbose "$($EXServer) Messages found: $($fndata.count)"
				Return $fndata
			} # Function Track-ExMessage

			$isdata += Track-ExMessage -EXServer $Using:EXServer -Sender $Using:Sender -Recipient $Using:Recipient -Start $Using:Start -End $Using:End -Subject $Using:Subject
write-verbose "$($Using:EXServer) Message count: $($isdata.count)"
			$isdata
			} # InlineScript
		} # foreach -parallel ($EXServer IN $EXServers)
write-verbose "Total Message count: $($exdata.count)"
	return $exdata
} # Workflow Main

# Makes an array of the internal Transport servers
Get-TransportServer | where {(test-connection $_.name -count 1 -quiet -erroraction silentlycontinue) -eq $true} | sort name | foreach {$EXServers+=$_.Name}
write-verbose "Server List: $EXServers"
write-verbose "Exchange Transport Seervers collected successfully: $? - Servers collected: $($EXServers.count)"

# Capture the output from Main
$trackingdata = Main -EXServers $EXServers -Sender $Sender -Recipient $Recipient -Start $Start -End $End -Subject $Subject

# Sort by TimeStamp and only select the fields we want
$trackingdata = $trackingdata | sort Timestamp | Select "Timestamp","ClientIp","ClientHostname","ServerIp","ServerHostname","ConnectorId","Source","EventId","InternalMessageId","MessageId","Recipients","RecipientStatus","TotalBytes","RecipientCount","RelatedRecipientAddress","MessageSubject","Sender","ReturnPath","MessageInfo","MessageLatency","MessageLatencyType","EventData"

# return the data differently based on what parameters were used
switch ($PsCmdlet.ParameterSetName)
    {
    "PSO"   { return $trackingdata; break}
    "GUI"   { $trackingdata | Out-GridView -Title "Message Tracking Summary"; break}
    default { $trackingdata  | Export-Csv -Path $FileName -NoTypeInformation -Force -Confirm:$false; break}
    } 
