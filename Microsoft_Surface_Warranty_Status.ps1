<#
Synopsis: A script that queries the Config Manager DB for all Microsoft Serial keys, runs them against the Microsoft warranty API, and exports everything to a CSV

Portions of the script (mainly the parts for connecting with Microsoft's Rest API) were ripped and modified from https://www.cyberdrain.com/automating-with-powershell-automating-warranty-information-reporting/

Made by BeholdenCypress in collaboration with Christopher Catlett (https://mdt2012.com/author/mdt2012/)
#>

$SQLSerials = Invoke-Sqlcmd -server SERVERNAME -Database ConfigManDB "SELECT
vWorkstationStatus.Name AS 'Computer name',
v_GS_COMPUTER_SYSTEM.Model0 AS 'Model',
v_GS_PC_BIOS.SerialNumber0 AS 'Serialnumber'
FROM
vWorkstationStatus INNER JOIN
v_GS_PC_BIOS ON vWorkstationStatus.ResourceID = v_GS_PC_BIOS.ResourceID INNER JOIN
v_GS_COMPUTER_SYSTEM ON v_GS_PC_BIOS.ResourceID = v_GS_COMPUTER_SYSTEM.ResourceID
WHERE
(vWorkstationStatus.OperatingSystem not like N'%server %') and (v_GS_COMPUTER_SYSTEM.Manufacturer0 like N'Microsoft%') and (v_GS_COMPUTER_SYSTEM.Model0 not like N'Virtual%')
Order by 'Computer name' ASC" | Select-Object -Property "Computer Name", "Model", "Serialnumber"

$outputpath = "c:\CSV\results.csv"
$results = [System.Collections.ArrayList]@()
foreach($device in $SQLSerials)
    {
    $body = ConvertTo-Json @{
        sku          = "Surface_"
        SerialNumber = $Device.Serialnumber
        ForceRefresh = $false
    }
    $today = Get-Date
    $PublicKey = Invoke-RestMethod -Uri 'https://surfacewarrantyservice.azurewebsites.net/api/key' -Method Get
    $AesCSP = New-Object System.Security.Cryptography.AesCryptoServiceProvider 
    $AesCSP.GenerateIV()
    $AesCSP.GenerateKey()
    $AESIVString = [System.Convert]::ToBase64String($AesCSP.IV)
    $AESKeyString = [System.Convert]::ToBase64String($AesCSP.Key)
    $AesKeyPair = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$AESIVString,$AESKeyString"))
    $bodybytes = [System.Text.Encoding]::UTF8.GetBytes($body)
    $bodyenc = [System.Convert]::ToBase64String($AesCSP.CreateEncryptor().TransformFinalBlock($bodybytes, 0, $bodybytes.Length))
    $RSA = New-Object System.Security.Cryptography.RSACryptoServiceProvider
    $RSA.ImportCspBlob([System.Convert]::FromBase64String($PublicKey))
    $EncKey = [System.Convert]::ToBase64String($rsa.Encrypt([System.Text.Encoding]::UTF8.GetBytes($AesKeyPair), $false))
    
    $FullBody = @{
        Data = $bodyenc
        Key  = $EncKey
    } | ConvertTo-Json
    
    $WarReq = Invoke-RestMethod -uri "https://surfacewarrantyservice.azurewebsites.net/api/v2/warranty" -Method POST -body $FullBody -ContentType "application/json"
    if ($WarReq.warranties) 
        {
            $WarObj = [PSCustomObject]@{
                'Serial'                = $SerialNumber
                'Warranty Product name' = $WarReq.warranties.name -join "`n"
                'StartDate'             = (($WarReq.warranties.effectivestartdate | sort-object -Descending | select-object -last 1) -split 'T')[0]
                'EndDate'               = (($WarReq.warranties.effectiveenddate | sort-object | select-object -last 1) -split 'T')[0]
                'Warranty Status'       = if (((($WarReq.warranties.effectiveenddate | sort-object | select-object -last 1) -split 'T')[0] | get-date) -le $today) { "Expired" } else { "OK" }
                'Client'                = $Client
            }
        }
    else
        {
            $WarObj = [PSCustomObject]@{
                'Serial'                = $SerialNumber
                'Warranty Product name' = 'Could not get warranty information'
                'StartDate'             = $null
                'EndDate'               = $null
                'Warranty Status'       = 'Could not get warranty information'
                'Client'                = $Client
            }
        }
    
    $info = $null
    $info = new-object -typename PSCustomObject
    
    $info | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Device.'Computer Name' -Force
    $info | Add-Member -MemberType NoteProperty -Name Model -Value $Device.Model -Force  
    $info | Add-Member -MemberType NoteProperty -Name Serial -Value $Device.Serialnumber -Force 
    $info | Add-Member -MemberType NoteProperty -Name StartDate -Value $WarObj.StartDate -Force
    $info | Add-Member -MemberType NoteProperty -Name EndDate -Value $WarObj.EndDate -Force
    $info | Add-Member -MemberType NoteProperty -Name WarrantyStatus -Value $WarObj.'Warranty Status' -Force
$results.Add($info)  
}
$results | export-csv -Path $outputpath -NoTypeInformation