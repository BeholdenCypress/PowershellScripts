#Written by: BeholdenCypress for Company Name
#Clear any stored errors
$error.clear()
#Get the current date
$CurrentDate = Get-Date -format "MMMM dd, yyyy"

#Start Log File
$VerbosePreference = "Continue"
$LogPath = "C:\Users\khowell\OneDrive - United Heritage Financial Group\Scripts\Testing\Logs\PasswordBot.log"
Get-ChildItem "$LogPath\*.log" | Where-Object LastWriteTime -LT (Get-Date).AddDays(-10) | Remove-Item -Confirm:$false
Start-Transcript -Path $LogPath -Append

# Get all users from AD, add them to a System.Array() using Username, Email Address, and Password Expiration date as a long date string and given the custom name of "PasswordExpiry"
try {
    $users = Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False -and PasswordLastSet -gt 0} -Properties "SamAccountName", "EmailAddress", "msDS-UserPasswordExpiryTimeComputed" | Select-Object -Property "SamAccountName", "EmailAddress", @{Name = "PasswordExpiry"; Expression = {[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} | Where-Object {$_.EmailAddress} -ErrorAction Stop
    $AuthDC = Get-WMIObject Win32_NTDomain | Select-Object "DomainControllerName" | Format-Table -HideTableHeaders | Out-String -ErrorAction Stop
    Write-Host 'AD Query Success'
    Write-Host "Machine authenticared to $($AuthDC)"
}
catch {
    Write-Host 'AD Query failed'
    $error
    Write-Host "Script failed to run on $($CurrentDate)"
    break
}

# Warning Date Variables
$FourteenDayWarnDate = (Get-Date).AddDays(14).ToLongDateString().ToUpper()
$TenDayWarnDate      = (Get-Date).AddDays(10).ToLongDateString().ToUpper()
$FiveDayWarnDate     = (Get-Date).AddDays(5).ToLongDateString().ToUpper()
$FourDayWarnDate     = (Get-Date).AddDays(4).ToLongDateString().ToUpper()
$ThreeDayWarnDate    = (Get-Date).AddDays(3).ToLongDateString().ToUpper()
$TwoDayWarnDate      = (Get-Date).AddDays(2).ToLongDateString().ToUpper()
$OneDayWarnDate      = (Get-Date).AddDays(1).ToLongDateString().ToUpper()

# Send-MailMessage parameters
$MailSender = 'United Heritage Password Bot <PasswordBot@companyname.com>'
$SMTPServer = 'emailrelay.companyname.com'

#Variables for image location
$Image1 = 'Path to Image 1.png'
$Image2 = 'Path to Image 2.png'
$Image3 = 'Path to Image 3.png'

#Convert images to base64 to embed into the HTML code.
$Image1AsBase64 = '<img src="data:image/png;base64,'+[convert]::ToBase64String((get-content $Image1 -encoding byte))+'">'
$Image2AsBase64 = '<img src="data:image/png;base64,'+[convert]::ToBase64String((get-content $Image2 -encoding byte))+'">'
$Image3AsBase64 = '<img src="data:image/png;base64,'+[convert]::ToBase64String((get-content $Image3 -encoding byte))+'">'

try {
    foreach($User in $Users) {
        $days = (([datetime]$User.PasswordExpiry) - (Get-Date)).days
    
        $WarnDate = Switch ($days) {
            14 {$FourteenDayWarnDate}
            10 {$TenDayWarnDate}
            5 {$FiveDayWarnDate}
            4 {$FourDayWarnDate}
            3 {$ThreeDayWarnDate}
            2 {$TwoDayWarnDate}
            1 {$OneDayWarnDate}
        }
    
        if ($days -in 14, 10, 5, 4, 3, 2, 1) {
            $SamAccount = $user.SamAccountName.ToUpper()
            $Subject    = "Windows Account Password for $($SamAccount) will expire soon"
            $EmailBody  = @"
                    <html> 
                        <body> 
                            <h1>Your Windows Account password will expire soon</h1> 
                            <H2>The Windows Account Password for <span style="color:red">$SamAccount</span> will expire in <span style="color:red">$days</span> days on <span style="color:red">$($WarnDate).</Span></H2>
                            <H3>If you need assistance changing your password, please reply to this email to submit a ticket</H3>
                            <br>
                            <br>
                            <br>
                            <br>
                            <br>
                            <H1>How to change your password</H1>
                            <p>While on any part of your computer, press the key combination <b>Ctrl + Alt + Del</b>. <br> This will bring you to the Windows Security screen. If you are working from a remote computer,<br>you will need to press <B>Ctrl + Alt + End</B> to get to the same screen</p>
                            <p>$Image1AsBase64</p>
                            <p>Click Change a Password. It will bring you to a new page. <br>Enter your old password, and then enter your new password twice.</p>
                            <p>$Image2AsBase64</p>
                            <p>Once you have confirmed your new password in both text boxes, either press enter or click the arrow to the right of the fourth text box. <br>If it completes successfully, it will say Your password has been changed.</p>
                            <p>$Image3AsBase64</p>
                            <p><img src="https://unitedheritage.com/images/logos/uh_web_logo_80.png"><img src="https://sublimityins.com/images/sic_logo_small.png"> </p>
                        </body> 
                    </html>
"@
            $MailSplat = @{
                To          = $User.EmailAddress
                From        = $MailSender
                SmtpServer  = $SMTPServer
                Subject     = $Subject
                BodyAsHTML  = $true
                Body        = $EmailBody
            }

            try {
                #Send-MailMessage @MailSplat -ErrorAction Stop
                Write-host "$($Days) days left, Email Sent to $($SamAccount)" -ErrorAction Stop
            }
            catch {
                Write-Host "Email failed to send"
                $error
                Write-Host "Script failed to run on $($CurrentDate)"
                Break
            }
        }
    }
}
catch {
    Write-Host 'Something Happened'
    $error
    Write-Host "Script failed to run on $($CurrentDate)"
    Break
}

if (!($error)) {
    Write-Host "The script ran successfully on $($CurrentDate)"
}
else {
    $error
}

Stop-Transcript
