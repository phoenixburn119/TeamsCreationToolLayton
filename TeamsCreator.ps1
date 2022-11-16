$global:TeamsName = ""
$global:TeamsID = ""
$global:TeamsAllInfo = ""
$ChannelList = Get-Content -Path "\\lcc-automation\AutomationPublic\Kinzer_Directory\TeamsCreationTool\ChannelList.txt"

Function PrintError($err) {
    #Used for return of standardized formated errors throughout the program.
    write-host $err -ForeGroundColor Red
}
Function MenuCreate {
    Write-Host "PROGRAM NAME PLACEHOLDER maybe Teams MultiTool cause I like it" -NoNewLine -ForeGroundColor Yellow
    Write-Host " "

    #Creates a main menu that detects userID's existance and also displays it. 
    if ($global:TeamsName) {
    Write-Host "Selected Team: $global:TeamsName"
    Write-Host "Debug:Selected Teams ID: $global:TeamsID"
    Write-Host "Debug:Selected TeamsAllInfo: $global:TeamsAllInfo"
    }
    Write-Host "Info: Info->Create 365 Group Instructions" -ForegroundColor Green
    Write-Host "1: Teams Selector"
    if ($global:TeamsID) {
        Write-Host "2: Create Channels And Permissions"
        Write-Host "3: Modify Permissions Of Channels !WIP!"
        Write-Host "4: Modify Permissions Of Team Owners/Members !WIP!"
    }
    Write-Host "Q: Press 'Q' to quit."
}
Function ModuleChecker {
    Write-Host "Module Checker Initiated" -ForegroundColor Blue
    If ((Get-InstalledModule -Name PowerShellGet).Name -eq "PowerShellGet") {
        Write-Host "PowerShellGet...Passed" -ForegroundColor Yellow
    }
    Else {
        Write-Host "Installing PowerShellGet Module..." -BackgroundColor Green
        Write-Host "Run in Admin Powershell -> Install-Module PowerShellGet -Force -AllowClobber" 
    }
    If ((Get-InstalledModule -Name MicrosoftTeams).Version -eq "4.2.0") {
        Write-Host "MicrosoftTeams 4.2.0...Passed" -ForegroundColor Yellow
    }
    Else {
        Write-Host "Installing MicrosoftTeams Module..." -BackgroundColor Green
        Write-Host "Run in Admin Powershell -> Install-Module MicrosoftTeams -AllowPrerelease -RequiredVersion "4.2.0"" 
        Pause
        Return
    }
    Write-Host "Continuing to program" -ForegroundColor White -BackgroundColor Green
    Start-Sleep -seconds 2
}
Function SignIn {
    Write-Host "Checking connection to MicrosoftTeams..." -BackgroundColor DarkGray
    Try{
        Get-CsOnlineUser -Identity jzinger2@laytonconstruction.com | Out-null
        Write-Host "Thanks for connecting to MicrosoftTeams prior ;)"
    } Catch{
        Write-Warning "Connecting to MicrosoftTeams...Please follow the popup"
        Connect-MicrosoftTeams
    }
}
Function TeamsSelector {
    SignIn
    $global:TeamsName = Read-Host -prompt "What is the name of the Team?"
    $global:TeamsAllInfo = Get-Team -DisplayName "$global:TeamsName"
    $global:TeamsID = $global:TeamsAllInfo.GroupID
    Write-Host "Displaying All Info..."
    $global:TeamsAllInfo
    $global:TeamsID
    Pause
    Return
}
Function 365GroupCreationMenu {
    Write-Host "Please create a Team on https://admin.microsoft.com/AdminPortal with the following specs."
    Write-Host ""
    Write-Host "1. Name it the project number and name (207416 CAPS Syringe Automation)"
    Write-Host "2. Add yourself as an owner"
    Write-Host "3. Email address is the Group name with no spaces (207416CAPSSyringeAutomation)"
    Write-Host "4. Make the privacy private"
    Write-Host "5. Leave the check for  Create a team for this group "
    Write-Host "6. Add the users to the team"
    Write-Host "7. Edit the permissions that the bottom two are the only checked boxes"
    Write-Host ""
    If ((Read-Host -prompt "Has this all been complete? (y/N)") -eq 'y' ) {
        Write-Host "Thats great!"
        Pause
    }
    Else {
        Return
    }
}
Function Get-TeamsChannelChecker($ChanName) {
    $TeamChannel = Get-TeamChannel -GroupId $global:TeamsID
    foreach ($ChannelSubject in $TeamChannel.DisplayName) {
        If($ChannelList.Contains($ChanName)) {
            Return $True
        }
    }
    Return $False
}
Function TeamsAddChannels {
    SignIn
    $UserName = Read-Host -Prompt "Enter users username (first.last)"
    If (!(Get-ADUser -Filter "sAMAccountName -eq '$UserName'")) { #Checks if the user exists in AD.
        Write-Host "User does not exist."
        Start-Sleep -Seconds 5
        Return
    }
    $UserName = $UserName + "@laytonconstruction.com"
    For($idx = 0; ($idx -lt $ChannelList.Count) -AND (Get-TeamsChannelChecker($ChannelList[$idx])); $idx++) {
        New-TeamChannel -GroupId $global:TeamsID -DisplayName $ChannelList[$idx] -MembershipType Private -Owner $UserName
        Start-Sleep -Seconds 2
    }
    Pause
}

Clear-Host
ModuleChecker
SignIn
#Main program loop.
While ($True) {
    Clear-Host
    MenuCreate
    $Selection = Read-Host "Please make a selection"
    switch ($Selection) {
        'info' { 365GroupCreationMenu }
        '1' { TeamsSelector }
        '2' { TeamsAddChannels }
        '3' {  }
        'q' {
            write-Host 'Hope we could help!' -ForegroundColor Blue
            If ((Read-Host -Prompt "Do you want to disconnect from TeamsOnline?(y/N)") -eq "y") {
                Disconnect-MicrosoftTeams
                Write-Host "Bye!"
                Exit
            }
        }
    }
}