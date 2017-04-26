# Here lie the global variables
$ScreenConnectDir = "C:\Program Files (x86)\ScreenConnect"
$SaveDirectory = "C:\Transcoded Sessions"
Add-Type -Path (Get-Item "$ScreenConnectDir\bin\ScreenConnect.Core.dll").FullName
Add-Type -Path (Get-Item "$ScreenConnectDir\bin\ScreenConnect.Server.dll").FullName
$TranscodeAttempts=0

# Returns true if a transcode has already happened (based on filename)
Function ExistingTranscode($fileNameToTest)
{
	# Returns true if the file exists
	Return (Test-Path "$SaveDirectory\$fileNameToTest")
}

# Transcodes a raw capture file to its output
Function Transcode
{
	Param(
		[Parameter(Mandatory=$false,Position=0)] $rawCapturePath,
		[Parameter(Mandatory=$false,Position=1)] $outputFileName
	)
	
	# If a transcode hasn't happened yet, do it
	If(-Not(ExistingTranscode($outputFileName)))
	{
	    Write-Host "Transcoding $outputFileName"
		
		Try
		{
			# Attempt the transcode
			$inputStream = [System.IO.File]::OpenRead($rawCapturePath)
			$toolkitInstance = [ScreenConnect.ServerToolkit]::Instance
			$encoder = $toolkitInstance.CreateVideoFileEncoder()
			[ScreenConnect.ServerExtensions]::Transcode($inputStream, $encoder, "$SaveDirectory\" + $outputFileName)
		}
		Catch
		{
			# Display encountered errors
			echo "Error processing $($captureFile): $($Error[0])"
		}
		Finally
		{
			# Close the source file reader
			$inputStream.Dispose()
		}
	}
	Else
	{
		# If the file exists, but it is less than 6000 bytes, redo it
		If((Get-Item ("$SaveDirectory\" + $outputFileName)).Length -lt 6000)
		{
		    If($TranscodeAttempts -lt 3)
			{
				Write-Host "Removing item... $outputFileName"
				Remove-Item ("$SaveDirectory\" + $outputFileName)
				Transcode -rawCapturePath $rawCapturePath -outputFileName $outputFileName
				$TranscodeAttempts += 1
			}
			Else
			{
				$TranscodeAttempts=0
				Break
			}
		}
		
	}
}

# Loop through the raw captures
$files = Get-ChildItem("$ScreenConnectDir\App_Data\Session\*")
ForEach ($captureFile in $files)
{
	$captureFileName = $captureFile.Name
	$captureFilePath = $captureFile.FullName
	
	# Split the filename into pieces (GUIDs and timestamp)
	$RegEx = "-"
	$GUIDS = $captureFileName -Split $RegEx
	
	# Obtain the session GUID
	$SessionGUID = ""
	For($i=0; $i -lt 5; $i++)
	{
		$SessionGUID = $SessionGUID + $GUIDS[$i] + "-"
	}
	$SessionGUID = $SessionGUID.Substring(0,$SessionGUID.Length-1)
	
	# Obtain the capture GUID
	$CaptureGUID = ""
	For($i=5; $i -lt 10; $i++)
	{
		$CaptureGUID = $CaptureGUID + $GUIDS[$i] + "-"
	}
	$CaptureGUID = $CaptureGUID.Substring(0,$CaptureGUID.Length-1)
	
	# Obtain the timestamp
	$timestamp = ""
	For($i=10; $i -lt 16; $i++)
	{
		$timestamp = $timestamp + $GUIDS[$i] + "-"
	}
	$timestamp = $timestamp.Substring(0,$timestamp.Length-1)
	
	# Format the filename as timestamp_SessionGUID_CaptureGUID.avi
	$outputFileName = $timestamp + "_" + $SessionGUID + "_" + $CaptureGUID + ".avi"
	
	Transcode -rawCapturePath $captureFilePath -outputFileName $outputFileName
}
