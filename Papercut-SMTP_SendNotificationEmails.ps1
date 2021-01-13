################################################################################
# Base script sourced from https://gist.github.com/ser1zw/4366363
# Modified by Ryan Kennedy
################################################################################


################################################################################
# Configure variables
$SMTPserver = "localhost"
$Port = "25"
$MailFrom = "from@example.com"
$RecipientTo = "to@example.com"
$EmailDirectory = "E:\tmp\Emails"
$KeepBackup = $true # if set to anything but $true the eml file will be deleted
################################################################################


Function SendCommand($stream, $writer, $command) {
  # Send command
	foreach ($line in $command) {
		$writer.WriteLine($line)
	}
	$writer.Flush()
	Start-Sleep -m 100

	# Get response
	$buff = New-Object System.Byte[] 4096
	$output = ""
	while ($True) {
		$size = $stream.Read($buff, 0, $buff.Length)
		if ($size -gt 0) {
			$output += $encoding.GetString($buff, 0, $size)
		}
		if (($size -lt $buff.Length) -or ($size -le 0)) {
			break
		}
	}

	if ([int]::Parse($output[0]) -gt 3) {
		throw $output
	}
	$output
}

Function SendMessage($SMTPserver, $Port, $MailFrom, $RecipientTo, $Email) {
	try {
		$socket = New-Object System.Net.Sockets.TcpClient($SMTPserver, $Port)
		$stream = $socket.GetStream()
		$stream.ReadTimeout = 1000
		$writer = New-Object System.IO.StreamWriter $stream
		$endOfMessage = "`r`n."
		SendCommand $stream $writer ("EHLO " + $SMTPserver)
		SendCommand $stream $writer ("MAIL FROM: <" + $MailFrom + ">")
		SendCommand $stream $writer ("RCPT TO: <" + $RecipientTo + ">")
		SendCommand $stream $writer "DATA"
		$content = (Get-Content $($Email.FullName)) -join "`r`n"
		SendCommand $stream $writer ($content + $endOfMessage)
		SendCommand $stream $writer "QUIT"
	}

	catch [Exception] {
		Write-Host $Error[0]
	}
	finally {
		if ($writer -ne $Null) {
			$writer.Close()
		}
		if ($socket -ne $Null) {
			$socket.Close()
		}
	}
}

$encoding = New-Object System.Text.AsciiEncoding
$Emails = Get-ChildItem $EmailDirectory -Filter "*.eml"

foreach ($Email in $Emails) {
    SendMessage $SMTPserver $Port $MailFrom $RecipientTo $Email | Out-Null
    if ($KeepBackup -eq $true) {
        Move-Item -Path $($Email.FullName) -Destination  "$EmailDirectory\Processed"
    } else {
        $Email | Remove-Item
    }
}
