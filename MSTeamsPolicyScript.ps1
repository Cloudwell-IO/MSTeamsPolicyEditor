###############################################
###############################################

###policyTypes should contain all the types of policies you'd like to manage within this script
$policyTypes = @("Meeting Policy","Messaging Policy","Live Events","App Permissions Policy","App Setup Policy","Call Park Policy","Calling Policy","Caller ID Policy","Teams Policy","Emergency Calling Policy","Emergency Call Routing Policy","Dial Plan","Voice Routing Policy")

###uncomment username and password to run script without credential prompt
#$username = "<username>"
#$password = "<password>"

#exports file log files to local script directory
$exportPath = (Get-Variable MyInvocation).Value.Path

#O365 domain name. This should be the primary login email domain. IE: microsoft.com
$loginDomain = "example.com"

###############################################
###############################################

$policyObjs = @()

foreach($policyType in $policyTypes)
{
    $policyObj = New-Object -TypeName psobject
    $policyObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $policyType
    switch($policyType){
        "Meeting Policy" {
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetPoliciesFunc" -Value {Get-CsTeamsMeetingPolicy}
            $policyObj | Add-Member -MemberType NoteProperty -Name "UserPropName" -Value TeamsMeetingPolicy
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetMembersFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyIdentity)
                process{
                    return (Get-CsOnlineUser | Where-Object {$_.TeamsMeetingPolicy -eq $policyIdentity} | Select-Object UserPrincipalName)
                }
            }
            $policyObj | Add-Member -MemberType ScriptMethod -Name "AddMemberFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyName,
                [Parameter(Mandatory=$true)][string]$userName,
                [Parameter(Mandatory=$true)][System.Windows.Forms.StatusBar]$statusBar
                )
                process{
                    $statusBar.Text = "Adding User..."
                    Grant-CsTeamsMeetingPolicy -PolicyName $policyName -Identity $userName
                    $statusBar.Text = "$($userName) successfully added to $($policyName)!"
                }
            }
        }
        "Calling Policy"{
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetPoliciesFunc" -Value {Get-CsTeamsCallingPolicy}
            $policyObj | Add-Member -MemberType NoteProperty -Name "UserPropName" -Value TeamsCallingPolicy
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetMembersFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyIdentity)
                process{
                    return (Get-CsOnlineUser | Where-Object {$_.TeamsCalingPolicy -eq $policyIdentity} | Select-Object UserPrincipalName)
                }
            }
            $policyObj | Add-Member -MemberType ScriptMethod -Name "AddMemberFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyName,
                [Parameter(Mandatory=$true)][string]$userName,
                [Parameter(Mandatory=$true)][System.Windows.Forms.StatusBar]$statusBar
                )
                process{
                    $statusBar.Text = "Adding User..."
                    Grant-CsTeamsCallingPolicy -PolicyName $policyName -Identity $userName
                    $statusBar.Text = "$($userName) successfully added to $($policyName)!"
                }
            }
        }
        "Messaging Policy"{
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetPoliciesFunc" -Value {Get-CsTeamsMessagingPolicy}
            $policyObj | Add-Member -MemberType NoteProperty -Name "UserPropName" -Value TeamsMessagingPolicy
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetMembersFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyIdentity)
                process{
                    return (Get-CsOnlineUser | Where-Object {$_.TeamsMessagingPolicy -eq $policyIdentity} | Select-Object UserPrincipalName)
                }
            }
            $policyObj | Add-Member -MemberType ScriptMethod -Name "AddMemberFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyName,
                [Parameter(Mandatory=$true)][string]$userName,
                [Parameter(Mandatory=$true)][System.Windows.Forms.StatusBar]$statusBar
                )
                process{
                    $statusBar.Text = "Adding User..."
                    Grant-CsTeamsMessagingPolicy -PolicyName $policyName -Identity $userName
                    $statusBar.Text = "$($userName) successfully added to $($policyName)!"
                }
            }
        }
        "App Permissions Policy"{
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetPoliciesFunc" -Value {Get-CsTeamsAppPermissionPolicy}
            $policyObj | Add-Member -MemberType NoteProperty -Name "UserPropName" -Value TeamsAppPermissionPolicy
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetMembersFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyIdentity)
                process{
                    return (Get-CsOnlineUser | Where-Object {$_.TeamsAppPermissionPolicy -eq $policyIdentity} | Select-Object UserPrincipalName)
                }
            }
            $policyObj | Add-Member -MemberType ScriptMethod -Name "AddMemberFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyName,
                [Parameter(Mandatory=$true)][string]$userName,
                [Parameter(Mandatory=$true)][System.Windows.Forms.StatusBar]$statusBar
                )
                process{
                    $statusBar.Text = "Adding User..."
                    Grant-CsTeamsAppPermissionPolicy -PolicyName $policyName -Identity $userName
                    $statusBar.Text = "$($userName) successfully added to $($policyName)!"
                }
            }
        }
        "App Setup Policy"{
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetPoliciesFunc" -Value {Get-CsTeamsAppSetupPolicy}
            $policyObj | Add-Member -MemberType NoteProperty -Name "UserPropName" -Value TeamsAppSetupPolicy
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetMembersFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyIdentity)
                process{
                    return (Get-CsOnlineUser | Where-Object {$_.TeamsAppSetupPolicy -eq $policyIdentity} | Select-Object UserPrincipalName)
                }
            }
            $policyObj | Add-Member -MemberType ScriptMethod -Name "AddMemberFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyName,
                [Parameter(Mandatory=$true)][string]$userName,
                [Parameter(Mandatory=$true)][System.Windows.Forms.StatusBar]$statusBar
                )
                process{
                    $statusBar.Text = "Adding User..."
                    Grant-CsTeamsAppSetupPolicy -PolicyName $policyName -Identity $userName
                    $statusBar.Text = "$($userName) successfully added to $($policyName)!"
                }
            }
        }
        "Call Park Policy"{
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetPoliciesFunc" -Value {Get-CsTeamsCallParkPolicy}
            $policyObj | Add-Member -MemberType NoteProperty -Name "UserPropName" -Value TeamsCallParkPolicy
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetMembersFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyIdentity)
                process{
                    return (Get-CsOnlineUser | Where-Object {$_.TeamsCallParkPolicy -eq $policyIdentity} | Select-Object UserPrincipalName)
                }
            }
            $policyObj | Add-Member -MemberType ScriptMethod -Name "AddMemberFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyName,
                [Parameter(Mandatory=$true)][string]$userName,
                [Parameter(Mandatory=$true)][System.Windows.Forms.StatusBar]$statusBar
                )
                process{
                    $statusBar.Text = "Adding User..."
                    Grant-CsTeamsCallParkPolicy -PolicyName $policyName -Identity $userName
                    $statusBar.Text = "$($userName) successfully added to $($policyName)!"
                }
            }
        }
        "Teams Policy"{
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetPoliciesFunc" -Value {Get-CsTeamsChannelsPolicy}
            $policyObj | Add-Member -MemberType NoteProperty -Name "UserPropName" -Value TeamsChannelsPolicy
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetMembersFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyIdentity)
                process{
                    return (Get-CsOnlineUser | Where-Object {$_.TeamsChannelsPolicy -eq $policyIdentity} | Select-Object UserPrincipalName)
                }
            }
            $policyObj | Add-Member -MemberType ScriptMethod -Name "AddMemberFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyName,
                [Parameter(Mandatory=$true)][string]$userName,
                [Parameter(Mandatory=$true)][System.Windows.Forms.StatusBar]$statusBar
                )
                process{
                    $statusBar.Text = "Adding User..."
                    Grant-CsTeamsChannelsPolicy -PolicyName $policyName -Identity $userName
                    $statusBar.Text = "$($userName) successfully added to $($policyName)!"
                }
            }
        }
        "Emergency Calling Policy"{
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetPoliciesFunc" -Value {Get-CsTeamsEmergencyCallingPolicy}
            $policyObj | Add-Member -MemberType NoteProperty -Name "UserPropName" -Value TeamsEmergencyCallingPolicy
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetMembersFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyIdentity)
                process{
                    return (Get-CsOnlineUser | Where-Object {$_.TeamsEmergencyCallingPolicy -eq $policyIdentity} | Select-Object UserPrincipalName)
                }
            }
            $policyObj | Add-Member -MemberType ScriptMethod -Name "AddMemberFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyName,
                [Parameter(Mandatory=$true)][string]$userName,
                [Parameter(Mandatory=$true)][System.Windows.Forms.StatusBar]$statusBar
                )
                process{
                    $statusBar.Text = "Adding User..."
                    Grant-CsTeamsEmergencyCallingPolicy -PolicyName $policyName -Identity $userName
                    $statusBar.Text = "$($userName) successfully added to $($policyName)!"
                }
            }
        }
        "Emergency Call Routing Policy"{
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetPoliciesFunc" -Value {Get-CsTeamsEmergencyCallRoutingPolicy}
            $policyObj | Add-Member -MemberType NoteProperty -Name "UserPropName" -Value TeamsEmergencyCallRoutingPolicy
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetMembersFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyIdentity)
                process{
                    return (Get-CsOnlineUser | Where-Object {$_.TeamsEmergencyCallRoutingPolicy -eq $policyIdentity} | Select-Object UserPrincipalName)
                }
            }
            $policyObj | Add-Member -MemberType ScriptMethod -Name "AddMemberFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyName,
                [Parameter(Mandatory=$true)][string]$userName,
                [Parameter(Mandatory=$true)][System.Windows.Forms.StatusBar]$statusBar
                )
                process{
                    $statusBar.Text = "Adding User..."
                    Grant-CsTeamsEmergencyCallRoutingPolicy -PolicyName $policyName -Identity $userName
                    $statusBar.Text = "$($userName) successfully added to $($policyName)!"
                }
            }
        }
        "Live Events Policy"{
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetPoliciesFunc" -Value {Get-CsTeamsMeetingBroadcastPolicy}
            $policyObj | Add-Member -MemberType NoteProperty -Name "UserPropName" -Value TeamsMeetingBroadcastPolicy
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetMembersFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyIdentity)
                process{
                    return (Get-CsOnlineUser | Where-Object {$_.TeamsMeetingBroadcastPolicy -eq $policyIdentity} | Select-Object UserPrincipalName)
                }
            }
            $policyObj | Add-Member -MemberType ScriptMethod -Name "AddMemberFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyName,
                [Parameter(Mandatory=$true)][string]$userName,
                [Parameter(Mandatory=$true)][System.Windows.Forms.StatusBar]$statusBar
                )
                process{
                    $statusBar.Text = "Adding User..."
                    Grant-CsTeamsMeetingBroadcastPolicy -PolicyName $policyName -Identity $userName
                    $statusBar.Text = "$($userName) successfully added to $($policyName)!"
                }
            }
        }
        "Caller ID Policy"{
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetPoliciesFunc" -Value {Get-CsCallingLineIdentity}
            $policyObj | Add-Member -MemberType NoteProperty -Name "UserPropName" -Value CallerIdPolicy
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetMembersFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyIdentity)
                process{
                    return (Get-CsOnlineUser | Where-Object {$_.CallerIdPolicy -eq $policyIdentity} | Select-Object UserPrincipalName)
                }
            }
            $policyObj | Add-Member -MemberType ScriptMethod -Name "AddMemberFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyName,
                [Parameter(Mandatory=$true)][string]$userName,
                [Parameter(Mandatory=$true)][System.Windows.Forms.StatusBar]$statusBar
                )
                process{
                    $statusBar.Text = "Adding User..."
                    Grant-CsCallingLineIdentity -PolicyName $policyName -Identity $userName
                    $statusBar.Text = "$($userName) successfully added to $($policyName)!"
                }
            }
        }
        "Dial Plan"{
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetPoliciesFunc" -Value {Get-CsTenantDialPlan}
            $policyObj | Add-Member -MemberType NoteProperty -Name "UserPropName" -Value DialPlan
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetMembersFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyIdentity)
                process{
                    return (Get-CsOnlineUser | Where-Object {$_.DialPlan -eq $policyIdentity} | Select-Object UserPrincipalName)
                }
            }
            $policyObj | Add-Member -MemberType ScriptMethod -Name "AddMemberFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyName,
                [Parameter(Mandatory=$true)][string]$userName,
                [Parameter(Mandatory=$true)][System.Windows.Forms.StatusBar]$statusBar
                )
                process{
                    $statusBar.Text = "Adding User..."
                    Grant-CsTenantDialPlan -PolicyName $policyName -Identity $userName
                    $statusBar.Text = "$($userName) successfully added to $($policyName)!"
                }
            }
        }
        "Voice Routing Policy"{
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetPoliciesFunc" -Value {Get-CsOnlineVoiceRoute}
            $policyObj | Add-Member -MemberType NoteProperty -Name "UserPropName" -Value OnlineVoiceRoutingPolicy
            $policyObj | Add-Member -MemberType ScriptMethod -Name "GetMembersFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyIdentity)
                process{
                    return (Get-CsOnlineUser | Where-Object {$_.OnlineVoiceRoutingPolicy -eq $policyIdentity} | Select-Object UserPrincipalName)
                }
            }
            $policyObj | Add-Member -MemberType ScriptMethod -Name "AddMemberFunc" -Value {
                [cmdletbinding()]
                param([Parameter(Mandatory=$true)]$policyName,
                [Parameter(Mandatory=$true)][string]$userName,
                [Parameter(Mandatory=$true)][System.Windows.Forms.StatusBar]$statusBar
                )
                process{
                    $statusBar.Text = "Adding User..."
                    Grant-CsOnlineVoiceRoute -PolicyName $policyName -Identity $userName
                    $statusBar.Text = "$($userName) successfully added to $($policyName)!"
                }
            }
        }
    }
    $policyObjs += $policyObj    
}


function getPolicyByType($type, $statusBar){
    
    $retVal = @()
    $policies = $null
    try{
        $policies = ($policyObjs | Where-Object {$_.Type -eq $type}).GetPoliciesFunc()
    }
    catch{
        $statusBar.Text = "Failure to get policy from policy object array."
    }
    if($null -ne $policies)
    {
        foreach($policy in $policies)
        {
            $retVal += $policy.Identity.Replace("Tag:","");
        }
    }
    return $retVal
}

function getPoliciesButtonClick($policyType, $listBox, $statusBar){
    $policies = getPolicyByType $policyType $statusBar
    $listBox.BeginUpdate()
    $listBox.Items.Clear()
    foreach($policy in $policies)
    {
        $listBox.Items.Add($policy);
    }
    $listBox.EndUpdate()
}

function getPolicyDetailsButtonClick($policyType, $policyIdentity, [System.Windows.Forms.TextBox]$textBox, $statusBar){
    $policy = $null
    $textBox.Clear()
    try{
        $policy = ($policyObjs | Where-Object {$_.Type -eq $policyType}).GetPoliciesFunc() | Where-Object {$_.Identity.Replace("Tag:","") -eq $policyIdentity}
    }
    catch{
        $statusBar.Text = "Failure to get policy object from policy object array."
    }
    if($null -ne $policy)
    {
        $textBox.Text = ($policy | Format-List | Out-String).Trim()
    }
}

function getPolicyMembersButtonClick($policyType, $policyIdentity, [System.Windows.Forms.ListBox]$listBox){
    $users = $null
    if($policyIdentity -eq "Global")
    {
        $policyIdentity = $null
    }
    $listBox.BeginUpdate()
    $listBox.Items.Clear()
    $listBox.Items.Add("Loading...")
    $listBox.EndUpdate()
    $listBox.BeginUpdate()
    $listBox.Items.Clear()
    if($policyIdentity -eq "Global")
    {
        $policyIdentity = $null
    }
    $users = ($policyObjs | Where-Object {$_.Type -eq $policyType}).GetMembersFunc($policyIdentity);
    if($null -ne $users)
    {
        $users = ($users | Sort-Object -Property UserPrincipalName)
        foreach($user in $users)
        {
            $listBox.Items.Add($user.UserPrincipalName)
        }
    }
    $listBox.EndUpdate()
}

function addRemoveMemberClick($users, $policyType, $policyName, [System.Windows.Forms.StatusBar]$statusBar){
    $userArr = $null
    if($users.GetType().Name -eq "SelectedObjectCollection")
    {
        $userArr = $users
    }
    else {
        $userArr = $users.Split(",")
    }
    foreach($user in $userArr)
    {
        if($user.IndexOf("@$($loginDomain)") -gt -1)
        {
            ($policyObjs | Where-Object {$_.Type -eq $policyType}).AddMemberFunc($policyName, $user, $statusBar);
        }
    }    
}

function userLookupClick($userName, $listBox){
    $listBox.BeginUpdate()
    $listBox.Items.Clear()
    $listBox.EndUpdate()
    $listBox.BeginUpdate()
    if($userName.IndexOf("@$($loginDomain)") -gt -1)
    {
        $user = Get-CsOnlineUser -Identity $username
        if($null -ne $user)
        {
            foreach($obj in $policyObjs)
            {
                $val = ($user | Select-Object $obj.UserPropName -ExpandProperty $obj.UserPropName)
                if($null -eq $val)
                {
                    $val = "Global"
                }
                $val = $val.ToString().Trim()
                $listBox.Items.Add("$($obj.Type): $($val)")
            }
        }
    }
    $listBox.EndUpdate()
}

function exportUserPolicyInfo([System.Windows.Forms.StatusBar]$statusBar){
    $statusBar.Text = "Processing export of all users and their policies. This will take a little time..."
    $membersArr = @()
    $path = "$($exportPath)AllUsersWithPolicies.csv"
    $users = Get-CsOnlineUser
    foreach($user in $users)
    {
        $memberObj = New-Object System.Object
        $memberObj | Add-Member -MemberType NoteProperty -Name "User" -Value $user.UserPrincipalName
        foreach($obj in $policyObjs)
        {
            $val = ($user | Select-Object $obj.UserPropName -ExpandProperty $obj.UserPropName)
            if($null -eq $val)
            {
                $val = "Global"
            }
            $val = $val.ToString().Trim()
            $memberObj | Add-Member -MemberType NoteProperty -Name $obj.Type -Value $val
        }
        $membersArr += $memberObj
    }
    $membersArr | Export-Csv -NoTypeInformation -Path $path
    $statusBar.Text = "Exported successfully to: $($path)"
}

function exportPolicyMembersClick($policyType, $policyName, $policyMembers, $allPolicies, [System.Windows.Forms.StatusBar]$statusBar){
    $membersArr = @()
    $path = "";
    if(!$allPolicies)
    {
        $statusBar.Text = "Processing export of $($policyName) members"
        $path = "$($exportPath)$($policyName)-Members.csv"
        foreach($member in $policyMembers)
        {
            $memberObj = New-Object System.Object
            $memberObj | Add-Member -MemberType NoteProperty -Name "Policy Type" -Value $policyType
            $memberObj | Add-Member -MemberType NoteProperty -Name "Policy Name" -Value $policyName
            $memberObj | Add-Member -MemberType NoteProperty -Name "User" -Value $member
            $membersArr += $memberObj
        }
        $membersArr | Export-Csv -NoTypeInformation -Path $path
    }
    else
    {
        $statusBar.Text = "Processing export of all policies' members. This will take a little time..."
        $path = "$($exportPath)$($policyType)-All Policies-Members.csv"
        $currPolicyObj = ($policyObjs | Where-Object {$_.Type -eq $policyType})
        $policies = getPolicyByType $policyType $statusBar
        foreach($policy in $policies)
        {
            $policyName = $policy
            if($policy -eq "Global")
            {
                $policy = $null
            }
            $users = $currPolicyObj.GetMembersFunc($policy);
            if($null -ne $users)
            {
                $users = ($users | Sort-Object -Property UserPrincipalName)
                foreach($user in $users)
                {
                    $memberObj = New-Object System.Object
                    $memberObj | Add-Member -MemberType NoteProperty -Name "Policy Type" -Value $policyType
                    $memberObj | Add-Member -MemberType NoteProperty -Name "Policy Name" -Value $policyName
                    $memberObj | Add-Member -MemberType NoteProperty -Name "User" -Value $user.UserPrincipalName
                    $membersArr += $memberObj
                }
            }
        }
        $membersArr | Export-Csv -NoTypeInformation -Path $path
    }
    $statusBar.Text = "Exported successfully to: $($path)"
}

function initializeForm(){
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $main_form                       = New-Object system.Windows.Forms.Form
    $main_form.ClientSize            = '1565,1080'
    $main_form.text                  = "Teams Policies Editor"
    $main_form.TopMost               = $false

    $policyInfoBox                   = New-Object system.Windows.Forms.Groupbox
    $policyInfoBox.height            = 997
    $policyInfoBox.width             = 1541
    $policyInfoBox.location          = New-Object System.Drawing.Point(7,38)

    $policyTypeCBLabel               = New-Object system.Windows.Forms.Label
    $policyTypeCBLabel.text          = "Policy Type"
    $policyTypeCBLabel.AutoSize      = $true
    $policyTypeCBLabel.width         = 25
    $policyTypeCBLabel.height        = 10
    $policyTypeCBLabel.location      = New-Object System.Drawing.Point(286,11)
    $policyTypeCBLabel.Font          = 'Microsoft Sans Serif,12'

    $policyTypeCB                    = New-Object system.Windows.Forms.ComboBox
    $policyTypeCB.width              = 252
    $policyTypeCB.height             = 25
    $policyTypes | ForEach-Object {[void] $policyTypeCB.Items.Add($_)}
    $policyTypeCB.location           = New-Object System.Drawing.Point(389,7)
    $policyTypeCB.Font               = 'Microsoft Sans Serif,12'

    $searchPolicyMembersButton       = New-Object system.Windows.Forms.Button
    $searchPolicyMembersButton.text  = "Search"
    $searchPolicyMembersButton.width  = 123
    $searchPolicyMembersButton.height  = 30
    $searchPolicyMembersButton.location  = New-Object System.Drawing.Point(1398,10)
    $searchPolicyMembersButton.Font  = 'Microsoft Sans Serif,8'

    $policyMembersListBox            = New-Object system.Windows.Forms.ListBox
    $policyMembersListBox.width      = 400
    $policyMembersListBox.height     = 400
    $policyMembersListBox.location   = New-Object System.Drawing.Point(993,35)
    $policyMembersListBox.SelectionMode = "MultiExtended"

    $policyDetailsTextbox            = New-Object system.Windows.Forms.TextBox
    $policyDetailsTextbox.multiline  = $true
    $policyDetailsTextbox.width      = 400
    $policyDetailsTextbox.height     = 400
    $policyDetailsTextbox.enabled    = $false
    $policyDetailsTextbox.location   = New-Object System.Drawing.Point(561,35)
    $policyDetailsTextbox.Font       = 'Microsoft Sans Serif,10'

    $policMembersLabel               = New-Object system.Windows.Forms.Label
    $policMembersLabel.text          = "Policy Members"
    $policMembersLabel.AutoSize      = $true
    $policMembersLabel.width         = 25
    $policMembersLabel.height        = 10
    $policMembersLabel.location      = New-Object System.Drawing.Point(994,13)
    $policMembersLabel.Font          = 'Microsoft Sans Serif,10'

    $searchPolicyMembersTextBox      = New-Object system.Windows.Forms.TextBox
    $searchPolicyMembersTextBox.multiline  = $false
    $searchPolicyMembersTextBox.width  = 285
    $searchPolicyMembersTextBox.height  = 20
    $searchPolicyMembersTextBox.location  = New-Object System.Drawing.Point(1097,10)
    $searchPolicyMembersTextBox.Font  = 'Microsoft Sans Serif,10'

    $selectPolicyTypeButton          = New-Object system.Windows.Forms.Button
    $selectPolicyTypeButton.text     = "Select"
    $selectPolicyTypeButton.width    = 123
    $selectPolicyTypeButton.height   = 30
    $selectPolicyTypeButton.location  = New-Object System.Drawing.Point(657,7)
    $selectPolicyTypeButton.Font     = 'Microsoft Sans Serif,10'

    $exportBox                       = New-Object system.Windows.Forms.Groupbox
    $exportBox.height                = 351
    $exportBox.width                 = 362
    $exportBox.text                  = "Export Information"
    $exportBox.location              = New-Object System.Drawing.Point(530,470)

    $policySettingsBox               = New-Object system.Windows.Forms.Groupbox
    $policySettingsBox.height        = 133
    $policySettingsBox.width         = 342
    $policySettingsBox.text          = "Policy Info Export"
    $policySettingsBox.location      = New-Object System.Drawing.Point(9,29)

    $allPoliciesExport               = New-Object system.Windows.Forms.CheckBox
    $allPoliciesExport.text          = "All"
    $allPoliciesExport.AutoSize      = $false
    $allPoliciesExport.width         = 95
    $allPoliciesExport.height        = 20
    $allPoliciesExport.location      = New-Object System.Drawing.Point(8,41)
    $allPoliciesExport.Font          = 'Microsoft Sans Serif,10'

    $selectedPolicyCheckbox          = New-Object system.Windows.Forms.CheckBox
    $selectedPolicyCheckbox.text     = "Selected (From Above)"
    $selectedPolicyCheckbox.AutoSize  = $false
    $selectedPolicyCheckbox.width    = 166
    $selectedPolicyCheckbox.height   = 20
    $selectedPolicyCheckbox.location  = New-Object System.Drawing.Point(8,62)
    $selectedPolicyCheckbox.Font     = 'Microsoft Sans Serif,10'

    $exportPolicySettingsButton      = New-Object system.Windows.Forms.Button
    $exportPolicySettingsButton.text  = "Export Policy Settings"
    $exportPolicySettingsButton.width  = 150
    $exportPolicySettingsButton.height  = 30
    $exportPolicySettingsButton.location  = New-Object System.Drawing.Point(182,11)
    $exportPolicySettingsButton.Font  = 'Microsoft Sans Serif,8'

    $exportPolicyMembersButton       = New-Object system.Windows.Forms.Button
    $exportPolicyMembersButton.text  = "Export Policy Members"
    $exportPolicyMembersButton.width  = 150
    $exportPolicyMembersButton.height  = 30
    $exportPolicyMembersButton.location  = New-Object System.Drawing.Point(182,51)
    $exportPolicyMembersButton.Font  = 'Microsoft Sans Serif,8'

    $addMembersBox                   = New-Object system.Windows.Forms.Groupbox
    $addMembersBox.height            = 531
    $addMembersBox.width             = 557
    $addMembersBox.text              = "Add Members"
    $addMembersBox.location          = New-Object System.Drawing.Point(978,459)

    $addMembersTextBox               = New-Object system.Windows.Forms.TextBox
    $addMembersTextBox.multiline     = $true
    $addMembersTextBox.width         = 400
    $addMembersTextBox.height        = 390
    $addMembersTextBox.enabled       = $true
    $addMembersTextBox.location      = New-Object System.Drawing.Point(17,55)
    $addMembersTextBox.Font          = 'Microsoft Sans Serif,10'

    $addMembersButton                = New-Object system.Windows.Forms.Button
    $addMembersButton.text           = "Add Member(s)"
    $addMembersButton.width          = 123
    $addMembersButton.height         = 30
    $addMembersButton.location       = New-Object System.Drawing.Point(426,55)
    $addMembersButton.Font           = 'Microsoft Sans Serif,8'

    $removeMemberButton              = New-Object system.Windows.Forms.Button
    $removeMemberButton.text         = "Remove Member(s)"
    $removeMemberButton.width        = 123
    $removeMemberButton.height       = 30
    $removeMemberButton.location     = New-Object System.Drawing.Point(1398,74)
    $removeMemberButton.Font         = 'Microsoft Sans Serif,8'

    $userLookupTextField             = New-Object system.Windows.Forms.TextBox
    $userLookupTextField.multiline   = $false
    $userLookupTextField.width       = 260
    $userLookupTextField.height      = 20
    $userLookupTextField.location    = New-Object System.Drawing.Point(13,19)
    $userLookupTextField.Font        = 'Microsoft Sans Serif,10'

    $userLookupButton                = New-Object system.Windows.Forms.Button
    $userLookupButton.text           = "Lookup"
    $userLookupButton.width          = 123
    $userLookupButton.height         = 30
    $userLookupButton.location       = New-Object System.Drawing.Point(284,19)
    $userLookupButton.Font           = 'Microsoft Sans Serif,8'

    $userLookupListBox               = New-Object system.Windows.Forms.ListBox
    $userLookupListBox.width         = 400
    $userLookupListBox.height        = 457
    $userLookupListBox.location      = New-Object System.Drawing.Point(7,60)

    $userLookupBox                   = New-Object system.Windows.Forms.Groupbox
    $userLookupBox.height            = 525
    $userLookupBox.width             = 440
    $userLookupBox.text              = "User Lookup"
    $userLookupBox.location          = New-Object System.Drawing.Point(11,467)

    $policyNamesListBox              = New-Object system.Windows.Forms.ListBox
    $policyNamesListBox.width        = 400
    $policyNamesListBox.height       = 400
    $policyNamesListBox.location     = New-Object System.Drawing.Point(11,39)

    $policyNamesLabel                = New-Object system.Windows.Forms.Label
    $policyNamesLabel.text           = "Policies"
    $policyNamesLabel.AutoSize       = $true
    $policyNamesLabel.width          = 25
    $policyNamesLabel.height         = 10
    $policyNamesLabel.location       = New-Object System.Drawing.Point(17,13)
    $policyNamesLabel.Font           = 'Microsoft Sans Serif,10'

    $policyDetailsLabel              = New-Object system.Windows.Forms.Label
    $policyDetailsLabel.text         = "Policy Details"
    $policyDetailsLabel.AutoSize     = $true
    $policyDetailsLabel.width        = 25
    $policyDetailsLabel.height       = 10
    $policyDetailsLabel.location     = New-Object System.Drawing.Point(565,13)
    $policyDetailsLabel.Font         = 'Microsoft Sans Serif,10'

    $policyDetailsButton             = New-Object system.Windows.Forms.Button
    $policyDetailsButton.text        = "Policy Details"
    $policyDetailsButton.width       = 123
    $policyDetailsButton.height      = 30
    $policyDetailsButton.location    = New-Object System.Drawing.Point(421,38)
    $policyDetailsButton.Font        = 'Microsoft Sans Serif,10'

    $policyMembersButton             = New-Object system.Windows.Forms.Button
    $policyMembersButton.text        = "Policy Members"
    $policyMembersButton.width       = 123
    $policyMembersButton.height      = 30
    $policyMembersButton.location    = New-Object System.Drawing.Point(421,77)
    $policyMembersButton.Font        = 'Microsoft Sans Serif,10'

    $selectedPolicyLabel             = New-Object system.Windows.Forms.Label
    $selectedPolicyLabel.text        = "Selected Policy:"
    $selectedPolicyLabel.AutoSize    = $true
    $selectedPolicyLabel.width       = 25
    $selectedPolicyLabel.height      = 10
    $selectedPolicyLabel.location    = New-Object System.Drawing.Point(17,453)
    $selectedPolicyLabel.Font        = 'Microsoft Sans Serif,10'

    $selectedPolicyValueTextBox      = New-Object system.Windows.Forms.TextBox
    $selectedPolicyValueTextBox.multiline  = $false
    $selectedPolicyValueTextBox.width  = 402
    $selectedPolicyValueTextBox.height  = 20
    $selectedPolicyValueTextBox.enabled  = $false
    $selectedPolicyValueTextBox.location  = New-Object System.Drawing.Point(17,475)
    $selectedPolicyValueTextBox.Font  = 'Microsoft Sans Serif,10'

    $addMembersInstructions          = New-Object system.Windows.Forms.Label
    $addMembersInstructions.text     = "Separate multiple users by commas (user1@$($loginDomain),user2@$($loginDomain)...)"
    $addMembersInstructions.AutoSize  = $true
    $addMembersInstructions.width    = 25
    $addMembersInstructions.height   = 10
    $addMembersInstructions.location  = New-Object System.Drawing.Point(17,26)
    $addMembersInstructions.Font     = 'Microsoft Sans Serif,10'

    $userInfoGroupBox                = New-Object system.Windows.Forms.Groupbox
    $userInfoGroupBox.height         = 144
    $userInfoGroupBox.width          = 337
    $userInfoGroupBox.text           = "User Info Export"
    $userInfoGroupBox.location       = New-Object System.Drawing.Point(14,185)

    $exportAllUsersPoliciesButton    = New-Object system.Windows.Forms.Button
    $exportAllUsersPoliciesButton.text  = "Export All Users`' Policies"
    $exportAllUsersPoliciesButton.width  = 150
    $exportAllUsersPoliciesButton.height  = 30
    $exportAllUsersPoliciesButton.location  = New-Object System.Drawing.Point(11,18)
    $exportAllUsersPoliciesButton.Font  = 'Microsoft Sans Serif,8'

    $main_form.controls.AddRange(@($policyInfoBox,$policyTypeCBLabel,$policyTypeCB,$selectPolicyTypeButton))
    $policyInfoBox.controls.AddRange(@($searchPolicyMembersButton,$policyMembersListBox,$policyDetailsTextbox,$policMembersLabel,$searchPolicyMembersTextBox,$exportBox,$addMembersBox,$removeMemberButton,$userLookupBox,$policyNamesListBox,$policyNamesLabel,$policyDetailsLabel,$policyDetailsButton,$policyMembersButton))
    $exportBox.controls.AddRange(@($policySettingsBox,$userInfoGroupBox))
    $policySettingsBox.controls.AddRange(@($allPoliciesExport,$selectedPolicyCheckbox,$exportPolicySettingsButton,$exportPolicyMembersButton))
    $addMembersBox.controls.AddRange(@($addMembersTextBox,$addMembersButton,$selectedPolicyLabel,$selectedPolicyValueTextBox,$addMembersInstructions))
    $userLookupBox.controls.AddRange(@($userLookupTextField,$userLookupButton,$userLookupListBox))
    $userInfoGroupBox.controls.AddRange(@($exportAllUsersPoliciesButton))
    $statusBar = New-Object System.Windows.Forms.StatusBar
    $main_form.Controls.Add($statusBar)

    $searchPolicyMembersButton.Add_Click({ searchMemberClick $this $_ })
    $selectPolicyTypeButton.Add_Click({ getPoliciesButtonClick $policyTypeCB.SelectedItem $policyNamesListBox $statusBar})
    $policyDetailsButton.Add_Click({ getPolicyDetailsButtonClick $policyTypeCB.SelectedItem $policyNamesListBox.SelectedItem $policyDetailsTextBox $statusBar})
    $policyMembersButton.Add_Click({
        $statusBar.Text = "Processing..."
        getPolicyMembersButtonClick $policyTypeCB.SelectedItem $policyNamesListBox.SelectedItem $policyMembersListBox
        $statusBar.Text = ""
        $exportPolicyMembersButton.enabled = $true
     })
    $exportPolicySettingsButton.Add_Click({ exportPolicySettingsClick $this $_ })
    $exportPolicyMembersButton.Add_Click({
        if($selectedPolicyCheckbox.Checked)
        {
            exportPolicyMembersClick $policyTypeCB.SelectedItem $policyNamesListBox.SelectedItem $policyMembersListBox.Items $false $statusBar
        } 
        if($allPoliciesExport.Checked)
        {
            exportPolicyMembersClick $policyTypeCB.SelectedItem $policyNamesListBox.SelectedItem $policyMembersListBox.Items $true $statusBar
        }
    })
    $exportAllUsersPoliciesButton.Add_Click({ exportUserPolicyInfo $statusBar})
    $userLookupButton.Add_Click({ userLookupClick $userLookupTextField.Text $userLookupListBox})
    $addMembersButton.Add_Click({ addRemoveMemberClick $addMembersTextBox.Text $policyTypeCB.SelectedItem $policyNamesListBox.SelectedItem $statusBar })
    $removeMemberButton.Add_Click({ 
        addRemoveMemberClick $policyMembersListBox.SelectedItems $policyTypeCB.SelectedItem $null $statusBar 
    })
    $policyNamesListBox.Add_SelectedIndexChanged({ 
        $selectedPolicyValueTextBox.Text = $this.SelectedItem
     })
    $allPoliciesExport.Add_CheckStateChanged({
        if($this.Checked)
        {
            $exportPolicyMembersButton.enabled = $true
        }
    })
    [void]$main_form.ShowDialog()
}

function initializeS4BPS($username, $pass){
    $creds = $null
    if($username -and $pass)
    {
        $pass = convertto-securestring -String $password -AsPlainText -Force
        $creds = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $pass
    }
    else
    {
        $creds = Get-Credential
    }
    if($null -ne $creds)
    {
        $sfbSession = New-CsOnlineSession -Credential $creds
        Import-PSSession $sfbSession
    }
    else {
        throw
    }
    
}

try{
    initializeS4BPS $username $password
    initializeForm
}
catch{
    Write-Host "Error connecting to S4B Powershell" -ForegroundColor Red
}
finally{
    Get-PSSession | Remove-PSSession
}


