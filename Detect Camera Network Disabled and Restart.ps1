$Adapter = Get-WmiObject -Class Win32_NetworkAdapter  -Filter "NetConnectionID LIKE 'CameraNetwork' and NetEnabled=false"
if($Adapter){
    c:
    cd '\Program Files\DevManView'
    .\DevManView.exe /disable_enable "Intel(R) Gigabit CT Desktop Adapter"
    $To         = "Warehouse Systems <REDACTED@overstock.com>";
    $From       = "Warehouse Systems <REDACTED@overstock.com>";
    $SmtpServer = "REDACTED.overstock.com";
    $Subject    = "REDACTED Network Adapter Enabled";
    $Results    = "The CameraNetwork network adapter on REDACTED has been enabled. If this continues to happen every 5 minutes, disable this Windows Scheduler Task and investigate further.";
    Send-MailMessage -smtpserver $SmtpServer -Priority High -To $To -From $From -Body $Results -Subject $Subject -BodyAsHtml
}