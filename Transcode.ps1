# Here lie the global variables
$ScreenConnectDir = "C:\Program Files (x86)\ScreenConnect"
$SaveDirectory = "C:\Transcoded Sessions"
Add-Type -Path (Get-Item "$ScreenConnectDir\bin\ScreenConnect.Core.dll").FullName
Add-Type -Path (Get-Item "$ScreenConnectDir\bin\ScreenConnect.Server.dll").FullName

# Returns true if a transcode has already happened (based on filename)
Function ExistingTranscode($fileNameToTest)
{
	# Get a list of already transcoded files
	$alreadyTranscoded = Get-ChildItem("$SaveDirectory\*")
	
	# Loop through the list of already transcoded filenames
	For($i=0; $i -lt ($alreadyTranscoded | Measure-Object).Count; $i++)
	{
		If($fileNameToTest -eq ($alreadyTranscoded[$i]).Name)
		{
			# The file already exists
			Return $True
		}
	}
	# The file doesn't exist
	Return $False;	
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
	
	# If a transcode hasn't happened yet, do it
	If(-Not(ExistingTranscode($outputFileName)))
	{
	    Write-Host "Transcoding $outputFileName"
		
		Try
		{
			# Attempt the transcode
			$inputStream = [System.IO.File]::OpenRead($captureFilePath)
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
		Write-Host "The file at $timestamp has already been transcoded."
	} 
}
