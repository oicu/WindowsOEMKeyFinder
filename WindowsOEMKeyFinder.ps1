<#
https://github.com/oicu
2025/12/10
Backup Windows Key, Windows Product Key Finder, Windows OEM Find Finder

修改执行策略以允许脚本运行:
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser

#>

function ConvertToKey {
    param(
        [byte[]]$Key
    )
    $Maps = "BCDFGHJKMPQRTVWXY2346789"

    # Check if OS is Windows 8
    $isWin8 = ([math]::Truncate($Key[14] / 6)) -band 1
    $Key[14] = ($Key[14] -band 0xF7) -bor (($isWin8 -band 2) * 4)

    $KeyOutput = ""
    $Last = 0

    for ($i = 24; $i -ge 0; $i--) {
        $Current = 0
        for ($j = 14; $j -ge 0; $j--) {
            $Current = $Current * 256
            $Current = $Key[$j] + $Current
            $Key[$j] = [math]::Truncate($Current / 24)
            $Current = $Current % 24
        }
        $KeyOutput = $Maps[$Current] + $KeyOutput
        $Last = $Current
    }

    if ($isWin8 -eq 1) {
        if ($Last -gt 0) {
            $keypart1 = $KeyOutput.Substring(1, $Last)
            $KeyOutput = $KeyOutput.Substring(0, 1) + $keypart1 + "N" + $KeyOutput.Substring(1 + $Last)
            $KeyOutput = $KeyOutput.Substring(1)
        } else {
            $KeyOutput = "N" + $KeyOutput.Substring(1)
        }
    }

    $formattedKey = $KeyOutput.Substring(0, 5) + "-" +
                    $KeyOutput.Substring(5, 5) + "-" +
                    $KeyOutput.Substring(10, 5) + "-" +
                    $KeyOutput.Substring(15, 5) + "-" +
                    $KeyOutput.Substring(20, 5)

    return $formattedKey
}

$regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
$dp  = Get-ItemPropertyValue -Path $regPath -Name "DigitalProductId"
#$dp4 = Get-ItemPropertyValue -Path $regPath -Name "DigitalProductId4"
$ProductName = Get-ItemPropertyValue -Path $regPath -Name "ProductName"
$ProductID = Get-ItemPropertyValue -Path $regPath -Name "ProductID"
$sls = Get-CimInstance -query 'select * from SoftwareLicensingService'

$ProductKey = ConvertToKey -Key $dp[52..67]
#$ProductKey = ConvertToKey -Key $dp4[808..823]
$OEMKey = $sls.OA3xOriginalProductKey
$OEMKeyDesc = $sls.OA3xOriginalProductKeyDescription
Write-Host ("=" * 78)
Write-Host "查看 Win7、Win8、Win10 使用的序列号、主板 OEM 序列号。`n"
Write-Host "Product Name: $ProductName"
Write-Host "Product ID: $ProductID"
Write-Host "Installed Key: $ProductKey`n"
Write-Host "OEM Key: $OEMKey"
Write-Host "Description: $OEMKeyDesc"
Write-Host ("=" * 78)
#Read-Host -Prompt "按下回车键以退出..."
Write-Host "按任意键继续..."
[System.Console]::ReadKey($true) | Out-Null
