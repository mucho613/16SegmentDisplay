$com = New-Object System.IO.Ports.SerialPort "COM4", 9600, ([System.IO.Ports.Parity]::None), 8, ([System.IO.Ports.StopBits]::One)
$com.Open()
$scrollWait = 100

while(1){
    $inputString = Read-Host "表示するテキストをMax9文字で入力してください(終了は""END"")"
    if($inputString -eq "END"){
        break
    }
    if($inputString.Length -lt 10){
        for($i = 0; $i -lt 9; $i++){
            $digitPosition = 0xe0 + $i
            $com.Write([Convert]::ToByte($digitPosition), 0, 1)
            $com.Write([Convert]::ToByte($inputString[$i]), 0, 1)
        }
    }
    else{
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