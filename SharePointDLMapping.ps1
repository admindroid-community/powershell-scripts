$URL = “https://your_domain.sharepoint.com/sites/test_site/Shared%20Documents” #<Replace with your document library URL copied in the first procedure> 
$IESession = Start-Process -file iexplore -arg $URL -PassThru -WindowStyle Hidden 
Sleep 20 
$IESession.Kill() 
$Network = new-object -ComObject WScript.Network 
$Network.MapNetworkDrive('Z:', $URL) #<Use the required drive name in place of ‘Z’> 