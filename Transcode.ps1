$ScreenConnectDir = "C:\Program Files (x86)\ScreenConnect"
Add-Type -Path (Get-Item "$ScreenConnectDir\bin\ScreenConnect.Core.dll").FullName
Add-Type -Path (Get-Item "$ScreenConnectDir\bin\ScreenConnect.Server.dll").FullName

foreach ($captureFile in Get-ChildItem("$ScreenConnectDir\App_Data\Session\*"))
{
    Try
    {
        #Write-Host "." -NoNewLine
        $inputStream = [System.IO.File]::OpenRead($captureFile.FullName)
		$toolkitInstance = [ScreenConnect.Toolkit]::Instance
		$toolkitInstance.CreateScreenEncoder()
		break
		#$encoder = $toolkitInstance.CreateVideoFileEncoder()
        #[ScreenConnect.ServerExtensions]::Transcode($inputStream, $encoder, $captureFile.FullName + ".avi")
    }
    Catch
    {
        echo "Error processing $($captureFile): $($Error[0])"
		Write-Host "Encountered error, stopping."
		break
    }
}
