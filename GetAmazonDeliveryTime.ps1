###
### First make sure you have some fresh items in cart
###

## Vars
$send_to = "real_verizon_mobile_number@vtext.com"
$send_from_password = "real_password"
$send_from_username = "real_gmail_username"
$check_page = "https://www.amazon.com/afx/slotselection/ref=ox_sc_fresh_slot_select%3Fclient=fullCart"

## Init
$hit = $false
$i = 0

## Loop Checks
while ($hit -eq $false) {
    ## Start a new page and open time slot selection
    Write-Progress -Activity "Starting Check" -PercentComplete 0
    $ie = new-object -ComObject "InternetExplorer.Application"
    sleep -Seconds 3
    Write-Progress -Activity "Open new IE" -PercentComplete 10
    $ie.navigate($check_page)
    sleep -Seconds 3
    while($ie.Busy) { Start-Sleep -Milliseconds 100 }

    try
    {
        ## Check for login needed (use browser saved password)
        if ($ie.LocationURL.ToLower() -ilike "*signin*")
        {
            $doc = $ie.Document
            $login_button = $doc.getElementsByTagName("Input")
            ## Needs to be figured out... for now on error just sign in... rest will error
            $login_button.click()
        }

        ## Get Page and Go To Cart
        $doc = $ie.Document
        Write-Progress -Activity "Going to Cart" -PercentComplete 25
        $cart = $doc.getElementById("nav-cart")
        $cart.click()
        sleep -Seconds 2
        while($ie.Busy) { Start-Sleep -Milliseconds 100 }
        $doc = $ie.Document
        Write-Progress -Activity "Going to Check Pages" -PercentComplete 50
        $ie.navigate($check_page)
        sleep -Seconds 3
        while($ie.Busy) { Start-Sleep -Milliseconds 100 }
        $doc = $ie.Document
        $day_buttons = $doc.getElementsByClassName("a-size-base-plus date-button-text a-text-bold")
        $day_buttons[$i].click()
        while($ie.Busy) { Start-Sleep -Milliseconds 100 }

        ## Check times
        $doc = $ie.Document
        Write-Progress -Activity "Checking Times" -PercentComplete 75
        $times = $doc.getElementsByName("slotsRadioGroup")
        $free_time = $times | ?{$_.isdisabled -eq $false}
        $hit = (($free_time | Measure-Object).count -ne 0)

        ## Check for time found
        if ($hit -ne $true){
            ## Up Day
            $i++    
            if ($i -ge 10){$i = 0}

            ## Sleep
            $random_sleep = (Get-Random -Minimum 15 -Maximum 120)
            for ($s = 1; $s -le $random_sleep; $s++)
            {
                Write-Progress -Activity "Sleeping" -Status "Sleeping for $s of $random_sleep seconds" -PercentComplete (($s/$random_sleep) * 100)
                sleep -Seconds 1
            }
            $ie.Quit()
        }
        else
        {
            Write-Progress -Activity "Hit Found" -PercentComplete 100
        }
    }
    catch
    {
        ## Something is wrong, alert, show current page
        Write-Progress -Activity "Error Found" -PercentComplete 100
        $sound = new-Object System.Media.SoundPlayer;
        $sound.SoundLocation="C:\Windows\media\notify.wav";
        $sound.Play();
        $ie.visible = $true
        sleep -Seconds 60
        $ie.Quit()
    }
}

## Hit found, show page
$ie.visible = $true
$free_time | %{$_.text}

## Send Email (Text)
$secpasswd = ConvertTo-SecureString $send_from_password -AsPlainText -Force
$mycreds = New-Object System.Management.Automation.PSCredential ($send_from_username, $secpasswd)
Send-MailMessage -To $send_to -Body "Amazon Delivery Available$([Environment]::NewLine)$($free_time | %{$_.text})$([Environment]::NewLine)$($check_page)" -From "$($send_from_username)@gmail.com" -SmtpServer "smtp.gmail.com" -Subject "Amazon Alert" -Port 587 -UseSsl -Credential $mycreds

## PLay Sound
$sound = new-Object System.Media.SoundPlayer;
$sound.SoundLocation="C:\Windows\media\Alarm01.wav";
while ($true) {$sound.Play(); sleep 5}