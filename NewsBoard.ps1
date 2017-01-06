$com = New-Object System.IO.Ports.SerialPort "COM4", 9600, ([System.IO.Ports.Parity]::None), 8, ([System.IO.Ports.StopBits]::One)
$com.Open()
$scrollWait = 100

$client = New-Object System.Net.WebClient

while(1){
    $content = $client.DownloadString("http://rss.nytimes.com/services/xml/rss/nyt/HomePage.xml")
    $xml = [xml]$content
    foreach($item in $xml.rss.channel.item){
        $inputString = $item.title.ToUpper()
        Write-Host $inputString
        $formattedString = "         " + $inputString + "         "
        for($i = 0; $i -le $inputString.Length + 9; $i++){
            $outputString = $formattedString.Substring($i, 9)
            for($j = 0; $j -lt 9; $j++){
                $digitPosition = 0xe0 + $j
                $com.Write([Convert]::ToByte($digitPosition), 0, 1)
                $com.Write([Convert]::ToByte($outputString[$j]), 0, 1)
            }
            Start-Sleep -m $scrollWait
        }
    }
}
$com.Close()
$com.Dispose()