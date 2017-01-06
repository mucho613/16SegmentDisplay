$com = New-Object System.IO.Ports.SerialPort "COM4", 9600, ([System.IO.Ports.Parity]::None), 8, ([System.IO.Ports.StopBits]::One)
$com.Open()
$spinnerState = 0
$spinnerWait = 5
$counter = 0
while(1){
    $time = Get-Date -Format "HHmmssff"

    # スピナーの状態遷移用
    $spinnerCharacter = [Convert]::ToByte("70", 16) + [Convert]::ToByte($spinnerState)
    $counter++
    if($counter -gt $spinnerWait){
        $counter = 0
        $spinnerState++
        if($spinnerState -gt 7){
            $spinnerState = 0
        }
    }

    # 9桁目(右端の桁)にスピナーを書き込む
    $com.Write([Convert]::ToByte("e8", 16), 0, 1)
    $com.Write([Convert]::ToByte($spinnerCharacter), 0, 1)

    # 数値の書き込み
    for($i = 0; $i -lt 8; $i++){
        if(($i -eq 1) -or ($i -eq 3) -or ($i -eq 5)){
            $digitPosition = 0xf0 + $i
        }
        else{
            $digitPosition = 0xe0 + $i
        }
        if(($i -eq 1) -or ($i -eq 3) -or ($i -eq 5)){
            $digitPosition = 0xf0 + $i
        }
        else{
            $digitPosition = 0xe0 + $i
        }
        $com.Write([Convert]::ToByte($digitPosition), 0, 1)
        $com.Write([Convert]::ToByte($time[$i]), 0, 1)
    }
    Start-Sleep -m 10
}
$com.Close()
$com.Dispose()