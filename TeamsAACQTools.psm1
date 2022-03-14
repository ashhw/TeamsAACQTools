class NasCQ {
    
    # Call Queue Name without the 'Optional' prefix      
    [string]$Name
    [String]$DisplayName

    #Custom input resource account name (What the end users will see when searching/calling in Teams)
    [string]$ResourceAccName

    [string]$ResourceAccountUPN

    [string]$CleanedRAName
    
    # Tenant domain in the following format: <YourTenant.onmicrosoft.com> or <domain.com>
    [ValidatePattern('^*.*')]
    [string]$TenantDomain

    # [OPTIONAL] Custom prefix for the call queue name
    [string]$Prefix

    # Call queue GUID populated once the call queue has been built
    [guid]$GUID

    # Call queue routing method required
    [ValidateSet('Attendant','Serial','RoundRobin','LongestIdle')]
    [string]$RoutingMethod = 'Attendant'

    # Call queue greeting music
    [string]$WelcomeMusicAudioFileId

    # Specify $True for default music on hold
    [ValidateSet('Y','N',$true,$false)]
    [string]$UseDefaultMusicOnHold

    # Agents to be added to the call queue
    [NasCQAgent[]]$Agents

    #Enable/disable presence based routing
    [ValidateSet('Y','N',$true,$false)]
    [string]$PresenceBasedRouting

    # Set $True or $False for agent opt out
    [ValidateSet('Y','N',$true,$false)]
    [string]$AllowOptOut

    # Agent alert time in seconds
    [int]$AgentAlertTime

    # Call queue overflow call limit
    [int]$OverflowThreshold

    # Call queue overflow action
    [ValidateSet('Disconnect','Forward','Voicemail','SharedVoicemail')]
    [string]$OverflowAction = 'Disconnect'

    # ObjectID of the overflow target e.g. UserA - GUID 01234-01234-123456-0000. Use the GUID
    [string]$OverflowActionTarget

    # Greeting to the caller when transfered to shared voicemail on overflow (Audio File ID)
    [string]$OverflowSharedVoicemailAudioFilePrompt

    # Greeting to the caller when transfered to shared voicemail on overflow (Text-to-Speech)
    [string]$OverflowSharedVoicemailTextToSpeechPrompt

    # Call queue imeout threshold in seconds
    [int]$TimeoutThreshold

    #Call queue timeout action required - Default 'Disconnect'
    [ValidateSet('Disconnect','Forward','Voicemail','SharedVoicemail')]
    [string]$TimeoutAction = 'Disconnect'

    # ObjectID of the timeout target e.g. UserA - GUID 01234-01234-123456-0000. Use the GUID
    [string]$TimeoutActionTarget

    # Greeting to the caller when transfered to shared voicemail on timeout (Audio File ID)
    [string]$TimeoutSharedVoicemailAudioFilePrompt

    # Greeting to the caller when transferred to shared voicemail on timeout (Text-to-Speech)
    [string]$TimeoutSharedVoicemailTextToSpeechPrompt

    # Id of the channel to connect a call queue to
    [string]$ChannelId

    # GUID of one of the owners of the team the channels belongs to
    [string]$ChannelUserObjectId

    # Enable/disable conference mode
    [ValidateSet('Y','N',$true,$false)]
    [string]$ConferenceMode

    # lets you add all the members of the distribution lists to the Call Queue. Use the DL GUID.
    [string[]]$DistributionLists

    # Turn on transcription for voicemails left by a caller on overflow
    [ValidateSet('Y','N',$true,$false)]
    [string]$EnableOverflowSharedVoicemailTranscription = 'N'

    # Turn on transcription for voicemails left by a caller on timeout
    [ValidateSet('Y','N',$true,$false)]
    [string]$EnableTimeoutSharedVoicemailTranscription = 'N'

    # Indicates the language that is used to play shared voicemail prompts
    [string]$LanguageID

    # Music to play when callers are placed on hold. This is the unique identifier of the audio file.
    [string]$MusicOnHoldAudioFileId

    # Resource account with phone number to the Call Queue for channel only
    [string[]]$OboResourceAccountIds

    #Build the display name from the custom prefix
    [NasCQ]PopulateDisplayNameFromPrefix(){
        $this.DisplayName = "{0}{1}" -f $this.Prefix, $this.Name
        return $this
    }

    [NasCQ]CleanDisplayName(){
        $this.DisplayName = Remove-StringSpecialCharacter -String $this.Name
        return $this
    }

    # [OPTIONAL] Specify the Music On Hold Audio File
    [string]$MusicOnHoldAudioFilePath

    # [OPTIONAL] Specify the Welcome Music Audio File
    [string]$WelcomeMusicAudioFilePath

    # Object ID(s) of the Resource Account associated with this call queue.
    [String[]]$ResourceAccount

    # List of phone numbers associated with this call queue
    [ValidatePattern('^\+.*')]
    [string[]]$PhoneNumber

}

class NasAA {
    
    [string]$Name
    [String]$LanguageID
    [string]$TimeZoneId

    [ValidatePattern('^*.*')]
    [string]$TenantDomain
    [string]$Prefix
    [guid]$GUID

    [PSCustomObject]$DefaultCallFlow
    [PSCustomObject]$CallFlows
    [PSCustomObject]$CallHandlingAssociations
    [switch]$EnableVoiceResponse
    [PSCustomObject]$ExclusionScope
    [string[]]$GreetingsSettingAuthorizedUsers
    [PSCustomObject]$InclusionScope
    [PSCustomObject]$Operator
    [string]$VoiceId

}

class NasCQAgent {
    [string]$AgentUPN
    #[string]$DisplayName
    [guid]$AgentGuid

    NasCQAgent(){}

    NasCQAgent([String]$AgentUPN,[guid]$AgentGuid){
        $this.AgentUPN = $AgentUPN
        $this.AgentGuid = $AgentGuid
    }
}

class NasObjLookup {
    [string]$TargetName
    [guid]$ObjGuid

    NasObjLookup(){}

    NasObjLookup([String]$TargetName,[guid]$ObjGuid){
        $this.TargetName = $TargetName
        $this.ObjGuid = $objGuid
    }
}

function Remove-StringSpecialCharacter {
    <#
.SYNOPSIS
    This function will remove the special character from a string.
.DESCRIPTION
    This function will remove the special character from a string.
    I'm using Unicode Regular Expressions with the following categories
    \p{L} : any kind of letter from any language.
    \p{Nd} : a digit zero through nine in any script except ideographic
    http://www.regular-expressions.info/unicode.html
    http://unicode.org/reports/tr18/
.PARAMETER String
    Specifies the String on which the special character will be removed
.PARAMETER SpecialCharacterToKeep
    Specifies the special character to keep in the output
.EXAMPLE
    Remove-StringSpecialCharacter -String "^&*@wow*(&(*&@"
    wow
.EXAMPLE
    Remove-StringSpecialCharacter -String "wow#@!`~)(\|?/}{-_=+*"
    wow
.EXAMPLE
    Remove-StringSpecialCharacter -String "wow#@!`~)(\|?/}{-_=+*" -SpecialCharacterToKeep "*","_","-"
    wow-_*
.NOTES
    Francois-Xavier Cat
    @lazywinadmin
    lazywinadmin.com
    github.com/lazywinadmin
#>
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [Alias('Text')]
        [System.String[]]$String,

        [Alias("Keep")]
        #[ValidateNotNullOrEmpty()]
        [String[]]$SpecialCharacterToKeep
    )
    PROCESS {
        try {
            IF ($PSBoundParameters["SpecialCharacterToKeep"]) {
                $Regex = "[^\p{L}\p{Nd}"
                Foreach ($Character in $SpecialCharacterToKeep) {
                    IF ($Character -eq "-") {
                        $Regex += "-"
                    }
                    else {
                        $Regex += [Regex]::Escape($Character)
                    }
                    #$Regex += "/$character"
                }

                $Regex += "]+"
            } #IF($PSBoundParameters["SpecialCharacterToKeep"])
            ELSE { $Regex = "[^\p{L}\p{Nd}]+" }

            FOREACH ($Str in $string) {
                Write-Verbose -Message "[INFO] Special character check - Original String: $Str"
                $Str -replace $regex, ""
            }
        }
        catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    } #PROCESS
}
Function Import-NasCQ {
    <#
    .SYNOPSIS
        Function to create the call queue object to pass into the New-NasTeamsCallQueue function.
    .DESCRIPTION
        This function will create an object to pass into the New-NasTeamsCallQueue function to create new call queues in Teams.
    .EXAMPLE
        PS C:\> New-NasCQ -Name "My Call Queue" -TenantDomain "tenant1.onmicrosoft.com" -Agents "User1@tenant.com","User2@tenant.com"
        This example will create a new call queue object with the name "My Call Queue" and the agents User1 and User 2 as members.
    .EXAMPLE
        PS C:\> New-NasCQ -Name "My Call Queue" -TenantDomain "tenant1.onmicrosoft.com" -Prefix "UC-CQ-UK-"

        This example will create a new call queue object with the name of "UC-CQ-UK-My Call Queue"

    .INPUTS
        None.
    .OUTPUTS
        None.
    .NOTES
    #>
    param (  
        # Call Queue Name without the 'Optional' prefix      
        #[string]$Name,

        # Tenant domain in the following format: <"YourTenant".onmicrosoft.com>
        [ValidatePattern('^*.*')]
        [string]$TenantDomain,

        # Call Queue data import
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [String]$CQData,

        #Give the user the option to install the required modules
        [switch]$InstallModules,

        #Create the call queues with no prompts
        [switch]$NoCreateCQ,

        [switch]$NoRA

    )

    #Define Transcript Log Files 
    #$logfile = (Get-Date).tostring("yyyyMMdd-hhmmss")
    #$transcriptfile = (New-Item -itemtype File -Path ".\" -Name ($logfile + ".log"))
    #Start-Transcript -Path $transcriptfile
    #Write-Host "Transcript logging started $($transcriptfile)"

    $errorStringPrefix = "[ERROR]"
    $InfoStringPrefix = "[INFO]"
    $RATypeAccountString = ":: RESOURCE ACCOUNT ::"

    if($InstallModules){
        Write-Verbose "$InfoStringPrefix Checking required modules are installed."

        #Check if the ImportExcel module is installed
        if (Get-InstalledModule -Name ImportExcel) {
            Write-Verbose "$InfoStringPrefix ImportExcel module exists, proceeding."
        } else {
            Write-Error "$errorStringPrefix ImportExcel module does not exist"
            Write-Host "To continue a module is required, would you like to install the ImportExcel module?"
            $Answer = Read-Host "Enter Y or N"
            if($Answer -eq 'Y'){
                Write-Host "Install-Module ImportExcel"
            }
            if($Answer -eq 'N'){
                Write-Verbose "$InfoStringPrefix Function stopped due to ImportExcel module requirement."
                break
            }
        }
    
        #Check if the MicrosoftTeams module is installed
        if (Get-InstalledModule -Name MicrosoftTeams) {
            Write-Verbose "$InfoStringPrefix MicrosoftTeams module exists, proceeding."
        } else {
            Write-Error "$errorStringPrefix MicrosoftTeams module does not exist"
            Write-Host "To continue a module is required, would you like to install the MicrosoftTeams module?"
            $Answer = Read-Host "Enter Y or N"
            if($Answer -eq 'Y'){
                Write-Host "Install-Module MicrosoftTeams"
            }
            if($Answer -eq 'N'){
                Write-Verbose "$InfoStringPrefix Function stopped due to MicrosoftTeams module requirement."
                break
            }
        }
    }

    Write-Verbose "$InfoStringPrefix Importing the Excel data from $CQData"
    #Import the Call Queue data from the Excel file
    $CQDataImport = Import-Excel -Path $CQData -WorksheetName "Call Queues"

    #Loop through and create the Call Queue objects
    ForEach($x in $CQDataImport) {

        $CQObj = [NasCQ]::new()

        #Need to clear these variables
        $WelcomeMusicAudioFileID,$MusicOnHoldAudioFileID,$DisplayName,$EnableOverflowSharedVoicemailTranscription = $null
        
        if($x.EnableOverflowSharedVoicemailTranscription -eq "Y"){
            $EnableOverflowSharedVoicemailTranscription = $x.EnableOverflowSharedVoicemailTranscription
        }else{
            $EnableOverflowSharedVoicemailTranscription = "N"
        }
        
        if($x.EnableTimeoutSharedVoicemailTranscription -eq "Y"){
            $EnableTimeoutSharedVoicemailTranscription = $x.EnableTimeoutSharedVoicemailTranscription
        }else{
            $EnableTimeoutSharedVoicemailTranscription = "N"
        }


        #Create the call queue objects from the Excel data input
        $CQObj.ResourceAccountUPN = $x.ResourceAccountUPN
        $CQObj.CleanedRAName = $x.CleanedName
        $CQObj.Name = $x.Name
        $CQObj.Prefix = $x.Prefix
        $CQObj.TenantDomain = $($TenantDomain)
        $CQObj.LanguageID = $x.LanguageID
        $CQObj.UseDefaultMusicOnHold = $x.UseDefaultMusicOnHold
        $CQObj.DistributionLists = $x.DistributionLists
        $CQObj.ChannelId = $x.ChannelId
        $CQObj.ChannelUserObjectId = $x.ChannelUserObjectId
        $CQObj.OboResourceAccountIds = $x.OboResourceAccountIds
        $CQObj.ConferenceMode = $x.ConferenceMode
        $CQObj.RoutingMethod = $x.RoutingMethod
        $CQObj.PresenceBasedRouting = $x.PresenceBasedRouting
        $CQObj.AllowOptOut = $x.AllowOptOut
        $CQObj.AgentAlertTime = $x.AgentAlertTime
        $CQObj.OverflowThreshold = $x.OverflowThreshold
        $CQObj.OverflowAction = $x.OverflowAction
        $CQObj.OverflowActionTarget = $x.OverflowActionTarget
        $CQObj.EnableOverflowSharedVoicemailTranscription = $EnableOverflowSharedVoicemailTranscription
        $CQObj.OverflowSharedVoicemailAudioFilePrompt = $x.OverflowSharedVoicemailAudioFilePrompt
        $CQObj.OverflowSharedVoicemailTextToSpeechPrompt = $x.OverflowSharedVoicemailTextToSpeechPrompt
        $CQObj.TimeoutThreshold = $x.TimeoutThreshold
        $CQObj.TimeoutAction = $x.TimeoutAction
        $CQObj.TimeoutActionTarget = $x.TimeoutActionTarget
        $CQObj.EnableTimeoutSharedVoicemailTranscription = $EnableTimeoutSharedVoicemailTranscription
        $CQObj.TimeoutSharedVoicemailAudioFilePrompt = $x.TimeoutSharedVoicemailAudioFilePrompt
        $CQObj.TimeoutSharedVoicemailTextToSpeechPrompt = $x.TimeoutSharedVoicemailTextToSpeechPrompt
        $CQObj.MusicOnHoldAudioFilePath = $x.MusicOnHoldAudioFilePath
        $CQObj.WelcomeMusicAudioFilePath = $x.WelcomeMusicAudioFilePath

        #Only populate the phone number if it exists otherwise it causes an error
        if($x.PhoneNumber){

            #Split multiple phone numbers for multiple resource accounts
            $CQObj.PhoneNumber = $x.PhoneNumber.split(",")
            Write-Verbose "$InfoStringPrefix Phone Numbers imported: $($CQObj.PhoneNumber)"
        }
        

        # Checking to see if we need to build this

        if(!($NoRA)){
            $ResourceAccount = $null
            $ResourceAccount = New-NasTeamsResourceAccount -CallQueue $CQObj
            #$ResourceAccount
        }

        #Create the call queues only if the parameter is specified, otherwise only import to the object
        if(!($NoCreateCQ)){
                if(!(Get-CsCallQueue -NameFilter "$($CQObj.Name)")){
                    Write-Verbose "$InfoStringPrefix $CQTypeAccountString Call Queue doesn't exist: $($x.Name). Creating the call queue."
                    #Call the New-NasTeamsCallQueue function to create the call queue
                    $CallQueue = $null
                    $CallQueue = New-NasTeamsCallQueue -CallQueue $CQObj 
                    if(!($NoRAAssociation)){
                        Write-Verbose "$InfoStringPrefix $CQTypeAccountString Call Queue: $($CQObj.Name) created, associating the resource account $($ResourceAccount.UserPrincipalName)"
                        $RAAssociation = $null
                        $RAAssociation = New-NasTeamsResourceAccountAssociation -CallQueue $CallQueue -ResourceAccountObjectID $ResourceAccount.ObjectID -ErrorAction Stop
                        Write-Verbose "$InfoStringPrefix $RATypeAccountString Resource account $($ResourceAccount.UserPrincipalName) has now been associated to $($CQObj.Name)"
                    }else{
                        Write-Verbose "$InfoStringPrefix $RATypeAccountString `$NoRAAssociation specified, skipping resource account association"
                    }
                    #Return Callqueue object
                    $CallQueue
                }else{
                    $CallQueue = $null
                    $CallQueue = Get-CsCallQueue -NameFilter "$($CQObj.Name)"
                    Write-Warning "$InfoStringPrefix $CQTypeAccountString Call Queue already exists: $($x.Name) exists as $($CQObj.Name). Checking resource account association."
                    if(!($NoRAAssociation)){
                        if(!(Get-CsOnlineApplicationInstanceAssociation -Identity $ResourceAccount.ObjectID -ErrorAction SilentlyContinue)){
                            Write-Verbose "$InfoStringPrefix $RATypeAccountString Resource account $($ResourceAccount.UserPrincipalName) not associated"
                            $RAAssociation = $null
                            $RAAssociation = New-NasTeamsResourceAccountAssociation -CallQueue $CallQueue -ResourceAccountObjectID $ResourceAccount.ObjectID -ErrorAction Stop
                            Write-Verbose "$InfoStringPrefix $RATypeAccountString Resource account $($ResourceAccount.UserPrincipalName) has now been associated to $($CQObj.Name)"
                        }else{
                            Write-Verbose "$InfoStringPrefix $CQTypeAccountString $($CQObj.Name) - Already associated with the following resource account $($ResourceAccount.UserPrincipalName)"
                        }
                    }else{
                        Write-Verbose "$InfoStringPrefix $RATypeAccountString `$NoRAAssociation specified, skipping resource account association"
                    }
                }
        }else{
            Write-Verbose "$InfoStringPrefix Call Queue imported to memory: $($CQObj.Name)"
            $CQObj
        }
    } # End import ForEach($x in $CQDataImport)

    #Stop-Transcript
    Write-Host "Call queue build completed, please refer to the transcript file for any errors."
}

function Get-NASTeamsLanguages {

    param (
        [Parameter(ValueFromPipeline)]
        [String]$rootFolder
    )

    $TeamsLanguages = [PSCustomObject]@{
        LanguageID = "ar-EG","ca-ES","da-DK","de-DE","en-AU","en-CA","en-GB","en-IN","en-US",
        "es-ES","es-MX","fi-FI","fr-CA","fr-FR","it-IT","ja-JP","ko-KR","nb-NO","nl-NL","pl-PL",
        "pt-PT","pt-BR","ru-RU","sv-SE","zh-CN","zh-HK","zh-TW","tr-TR","cs-CZ","th-TH","el-GR",
        "hu-HU","sk-SK","hr-HR","sl-SI","id-ID","ro-RO","vi-VN"
        Name = "Arabic (Egypt)","Catalan (Catalan)","Danish (Denmark)","German (Germany)","English (Australia)",
        "English (Canada)","English (United Kingdom)","English (India)","English (United States)","Spanish (Spain)",
        "Spanish (Mexico)","Finnish (Finland)","French (Canada)","French (France)","Italian (Italy)","Japanese (Japan)",
        "Korean (Korea)","Norwegian, BokmÃ¥l (Norway)","Dutch (Netherlands)","Polish (Poland)","Portuguese (Portugal)",
        "Portuguese (Brazil)","Russian (Russia)","Swedish (Sweden)","Chinese (Simplified, PRC)","Chinese (Traditional, Hong Kong S.A.R.)",
        "Chinese (Traditional, Taiwan)","Turkish (Turkey)","Turkish (Turkey)","Thai (Thai)","Greek (Greek)","Hungarian (Hungary)",
        "Slovak (Slovakia)","Croatian (Croatia)","Slovenian (Slovenia)","Indonesian (Indonesia)","Romanian (Romania)","Vietnamese (Vietnam)"
    }

    $TeamsLanguages | Sort-Object LanguageID | Export-Excel -Path "$rootFolder\AACQDataImport.xlsx" -WorksheetName "Languages" -NoNumberConversion "Name" -BoldTopRow -AutoSize
}

function Get-NASAgentGuid {
    <#
    .SYNOPSIS
        Synopsis
    .DESCRIPTION
        Description here.
    .EXAMPLE
        PS C:\> 
        This example 
    .EXAMPLE
        Example 2 here
    .INPUTS
        None.
    .OUTPUTS
        None.
    .NOTES
    #>
    [CmdletBinding()]
    param (
        # Parameter help description
        [Parameter(Mandatory=$True,ValueFromPipeline)]
        [string[]]$AgentUPN
    
    )
    Begin {Write-Verbose "Looking for Call Queue agents to build the agent objects."}
    process {
        ForEach-Object {
            try{ 
                $TypeObj = (Get-CsOnlineUser -Identity $_ | Get-Member)[0].typename
                Write-Verbose "Raw Data = $TypeObj"
                #Try and get the user
                if($TypeObj -like "*UserMas"){
                    $CQAgentGUID = (Get-CsOnlineUser -Identity $_).Identity
                    Write-Verbose "Converted New way: $CQAgentGUID"
                }else{
                    $CQAgentGUID = (Get-CsOnlineUser -Identity $_).id.split(",")[0].split("=")[1]
                    Write-Verbose "Converted Old way: $CQAgentGUID"
                }
                
            } Catch {
                Throw "Unable to find user $AgentUPN"
            }
            Write-Verbose "Looking for Call Queue agent: $AgentUPN - ObjectID: $($CQAgentGUID)"
            
            #Create the NasCQAgent object with the agents UPN and objectID
            [NasCQAgent]::new($AgentUPN,$CQAgentGUID)
            
            Write-Verbose "Found Call Queue agent: $AgentUPN - ObjectID: $($CQAgentGUID)"
        }
    }
    End{Write-Verbose "Found all call queue agents."}
}

function Get-NASObjectGuid {
    <#
    .SYNOPSIS
        Synopsis
    .DESCRIPTION
        Description here.
    .EXAMPLE
        PS C:\> 
        This example 
    .EXAMPLE
        Example 2 here
    .INPUTS
        None.
    .OUTPUTS
        None.
    .NOTES
    #>
    [CmdletBinding()]
    param (
        # Parameter help description
        [Parameter(Mandatory=$True,ValueFromPipeline)]
        [string]$TargetName
    
    )
    Begin {
        Write-Verbose "Looking for object GUIDs to set Overflow/Timeout action"
        $foundInvalid = $false
    }
    process {
        ForEach-Object {
            try{ 
                $TypeObj = (Get-CsOnlineUser -Identity $_ | Get-Member)[0].typename
                Write-Verbose "Raw Data = $TypeObj"
                #Try and get the user
                if($TypeObj -like "*UserMas"){
                    $objGuid = (Get-CsOnlineUser -Identity $_).Identity
                    Write-Verbose "Converted New way: $objGuid"
                }else{
                    $objGuid = (Get-CsOnlineUser -Identity $_).id.split(",")[0].split("=")[1]
                    Write-Verbose "Converted Old way: $objGuid"
                }
                
            } Catch {
                Write-Error "Unable to find object $($TargetName)"
                $foundInvalid = $true
            }
            if($objGuid){
                Write-Verbose "Building the object: $($TargetName) - ObjectID: $($objGuid)"
                #Create the NasCQAgent object with the agents UPN and objectID
                [NasObjLookup]::new($TargetName,$objGuid)
                Write-Verbose "Object built: $($TargetName) - ObjectID: $($objGuid)"
            }else{
                Write-Error "Object doesn't exist $($TargetName)"
            }
        }
    }
    End{
        if(!($foundInvalid)){ Write-verbose "Found all objects."}
    }
}

function New-NasTeamsResourceAccount {

    param(
        [Parameter(Mandatory,ParameterSetName = 'CallQueue')]
        [NasCQ]$CallQueue,

        [Parameter(Mandatory,ParameterSetName = 'AutoAttendant')]
        [NasAA]$AutoAttendant

    )

    if($CallQueue){
        $AppID = "11cd3e2e-fccb-42ad-ad00-878b93575e07"
        #$Prefix = "racq-cc-lll-"
        $DisplayName = $CallQueue.CleanedRAName
        #$TenantDomain = $CallQueue.TenantDomain
        $PhoneNumber = $CallQueue.PhoneNumber
    }else{
        $AppID = "ce933385-9390-45d1-9512-c8d228074e07"
        #$Prefix = "raaa-cc-lll-"
        $DisplayName = $AutoAttendant.CleanedRAName
        #$TenantDomain = $AutoAttendant.TenantDomain
        $PhoneNumber = $AutoAttendant.PhoneNumber
    }
    
    #Define the error verbose
    $errorStringPrefix = "[ERROR]"
    $InfoStringPrefix = "[INFO]"
    $RATypeAccountString = ":: RESOURCE ACCOUNT ::"

    $i = 0
    $RAAccountUPN = $null
    # we are doing this to prevent an error in foreach when the user hasn't provided an phone number, so we create a dummy value
    if(!($PhoneNumber)){$PhoneNumber = ""}
    $PhoneNumber | ForEach-Object {
        $i++
        if($PhoneNumber.count -eq 1){
            # Set to null so that single resource accounts don't have a numbered suffix
            $i = $null 
        }
        
        #Lets build the resource account UPN
        #if($CallQueue.Prefix){
        #    $RAAccountUPN = "{0}{1}{2}@{3}" -f $($CallQueue.Prefix), $($DisplayName), $i, $($TenantDomain.Replace(" ",""))
        #}else{
        #    $RAAccountUPN = "{0}{1}{2}@{3}" -f $Prefix, $DisplayName, $i, $($TenantDomain.Replace(" ",""))
        #    # racq-cc-lll-ITsupport1@wardmanor.onmicrosoft.com
        #    # racq-cc-lll-ITsupport2@wardmanor.onmicrosoft.com
        #}

        $RAAccountUPN = $CallQueue.ResourceAccountUPN

        #(Get-CsCallQueue -WarningAction SilentlyContinue).where({$_.identity -ne '0362bd89-d0d5-41a2-9564-8a184c8a084f'}) | % { $_.ApplicationInstances | % { Remove-CsOnlineApplicationInstanceAssociation -Identities  $_ }; $_ | Remove-CsCallQueue }

        $NewRA = Get-CsOnlineApplicationInstance -Identity $RAAccountUPN -ErrorAction SilentlyContinue
        if(!($NewRA)){
                Write-Verbose "$($InfoStringPrefix) $($RATypeAccountString) $($RAAccountUPN) - Account doesn't exist, moving to creation."
                # Create resource account of call queue type
                $RAParameters = @{
                    UserPrincipalName = $RAAccountUPN
                    ApplicationId = $AppID
                    DisplayName = "$DisplayName"
                }
                $NewRA = New-CsOnlineApplicationInstance @RAParameters
                Write-Verbose "$InfoStringPrefix $RATypeAccountString $($RAAccountUPN) - Account has now been created."

                #Write-Verbose "$InfoStringPrefix $RATypeAccountString $($RAAccountUPN) - Account is available for use, moving to call queue checks."
                if($CallQueue){
                    $CallQueue.ResourceAccount += $NewRA.objectID
                }

        #The Resource Account exists, skip to creating the call queue    
        }Else{
            #$NewRA = Get-CsOnlineApplicationInstance -Identity $RAAccountUPN
            Write-Warning "$InfoStringPrefix $RATypeAccountString $($RAAccountUPN) - Account already exists, skipping to call queue creation."
            $CallQueue.ResourceAccount += $NewRA.objectID
        }

        $NewRA
        Get-PSSession | Where-Object name -like "SfBPowerShellSessionViaTeamsModule*" | Remove-PSSession 
        Write-Verbose "Clearing down PS Sessions to avoid session congestion"
    }
}

function New-NasTeamsResourceAccountAssociation{

    [CmdletBinding()]
    param(
        [Parameter(Mandatory,ParameterSetName = 'CallQueue')]
        $CallQueue,

        [Parameter(Mandatory,ParameterSetName = 'AutoAttendant')]
        [NasAA]$AutoAttendant,

        [parameter(Mandatory)]
        [string]$ResourceAccountObjectID
    )

    if($CallQueue){
        $ConfigurationType = 'CallQueue'
    }else{
        $ConfigurationType = 'AutoAttendant'
    }

    #PS7 only
    #$ConfigurationType = ($CallQueue) ? 'CallQueue' : 'AutoAttendant'

    $i = $null
    $i = 0
    while (!(Get-CsOnlineApplicationInstance -Identities $ResourceAccountObjectID)) {
        Write-Verbose "Resource account: $ResourceAccountObjectID is not ready yet"
        if($i -gt 5){
            Write-Error "ERROR ::: Resource account: $ResourceAccountObjectID is not available for use yet. Please try again later."
            break
        }else{
            Write-Verbose "Looping back round to check if the resource account is ready."
        }
        $i++
    }
    Write-Verbose "Resource account: $ResourceAccountObjectID available, moving on."

    if(($CallQueue | gm)[0].typename -eq "NasCQ"){
        $RAConfigurationID = $CallQueue.Guid.Guid
    }else{
        $RAConfigurationID = $CallQueue.Identity
    }

    if(Get-CsOnlineApplicationInstance -Identities $ResourceAccountObjectID){
        $NewCQAppInstanceParameters = @{
            Identities = $ResourceAccountObjectID
            ConfigurationId = $RAConfigurationID
            ConfigurationType = $ConfigurationType
            ErrorAction = 'Stop'
        }
        New-CsOnlineApplicationInstanceAssociation @NewCQAppInstanceParameters -ErrorAction Stop
    }else{
        Write-Error "Resource account: $ResourceAccountObjectID cannot be associated. Terminating script."
        break
    }

}

function New-NasTeamsCallQueue {
        <#
    .SYNOPSIS
        Synopsis
    .DESCRIPTION
        Description here.
    .EXAMPLE
        PS C:\> 
        This example 
    .EXAMPLE
        Example 2 here
    .INPUTS
        None.
    .OUTPUTS
        None.
    .NOTES
    #>
    [CmdletBinding()]
    param (
        #[Parameter(Mandatory=$True)]
        [NasCQ]$CallQueue,
        #[switch]$NoResourceAccount,
        [switch]$force
    )

    #Variables
    #Define the error verbose
    $errorStringPrefix = "[ERROR]"
    $InfoStringPrefix = "[INFO]"
    $CQTypeAccountString = ":: CALL QUEUE ::"
    #$CustomCQSuffix = "CQ"

    #Path used for music based on excel file location
    $rootPath = $CQData | Split-Path -Parent

    Write-Verbose "$InfoStringPrefix $CQTypeAccountString $($CallQueue.Name) - Checking if the call queue exists."

    #Check if the call queue already exists, if it doesn't, then create the call queue
    if(!(Get-CsCallQueue -NameFilter "$($CallQueue.Name)")){

        Write-Verbose "$InfoStringPrefix $CQTypeAccountString $($CallQueue.Name) - Call queue doesn't exist, moving to creation."

        ## Audio file checks and uploads
        #Check if the WelcomeMusic has been specified in the data.
        if($CallQueue.WelcomeMusicAudioFilePath){
            Write-Verbose "$InfoStringPrefix $CQTypeAccountString $($CallQueue.Name) :: WelcomeMusic audio file found, checking path $rootPath$($CallQueue.WelcomeMusicAudioFilePath) exists."
            # Test the path, ensure it can be reached
            try {
                $PathTest = Test-Path -Path $rootPath$($CallQueue.WelcomeMusicAudioFilePath) -ErrorAction Stop
            }
            catch {
                throw "$errorStringPrefix $CQTypeAccountString $($CallQueue.Name) :: Path error: $rootPath$($CallQueue.WelcomeMusicAudioFilePath) not found/unreachable."
            }

            # Path returns true, therefore import and upload the file
            if($PathTest){
                #Generate random filename for audio file
                $WelcomeMusicFilename = "WelcomeMusic" + $CallQueue.Name + (Get-Random -Minimum 1 -Maximum 1000) + ".mp3"

                #Lets create duplicate file
                #Build the filename path
                $ParentPath = "$rootPath$($CallQueue.WelcomeMusicAudioFilePath)" | Split-Path -Parent
                $WelcomeMusicCopyFilename = "$ParentPath\$WelcomeMusicFilename"
                Copy-Item -Path "$rootPath$($CallQueue.WelcomeMusicAudioFilePath)" -Destination $WelcomeMusicCopyFilename

                Write-Verbose "$($WelcomeMusicCopyFilename)"

                #Get the audio file from the path
                $WelcomeMusicAudioFileContent = Get-Content -Path $WelcomeMusicCopyFilename -AsByteStream -ReadCount 0
                Write-Verbose "$InfoStringPrefix $CQTypeAccountString $($CallQueue.Name) :: Path exists: WelcomeMusic audio file $rootPath$($CallQueue.WelcomeMusicAudioFilePath)"
                Write-Verbose "$InfoStringPrefix $CQTypeAccountString $($CallQueue.Name) :: WelcomeMusic audio file $rootPath$($CallQueue.WelcomeMusicAudioFilePath) imported."

                # Import the audio file into Teams
                $WelcomeMusicAudioFile = Import-CsOnlineAudioFile -ApplicationId "HuntGroup" -FileName $WelcomeMusicFilename -Content $WelcomeMusicAudioFileContent  

                # Grab the ID - Required for the call queue
                $WelcomeMusicAudioFileID = $WelcomeMusicAudioFile.id

                Write-Verbose "Removing audio file copy: $WelcomeMusicCopyFilename"
                Remove-Item -Path $WelcomeMusicCopyFilename

                Write-Verbose "$InfoStringPrefix $CQTypeAccountString $($CallQueue.Name) :: WelcomeMusic audio file uploaded to Teams. ID: $($WelcomeMusicAudioFileID)"
            } else {
                Write-Error "$errorStringPrefix $CQTypeAccountString $($CallQueue.Name) :: WelcomeMusic file path $rootPath$($CallQueue.WelcomeMusicAudioFilePath) unreachable."
            }

        }

        #Check if the MusicOnHold has been specified in the data.
        if($CallQueue.MusicOnHoldAudioFilePath){
            Write-Verbose "$InfoStringPrefix $CQTypeAccountString $($CallQueue.Name) :: MusicOnHold audio file found, checking path $rootPath$($CallQueue.MusicOnHoldAudioFilePath) exists."
            # Test the path, ensure it can be reached
            try {
                $PathTest = Test-Path -Path "$rootPath$($CallQueue.MusicOnHoldAudioFilePath)" -ErrorAction Stop
                Write-Verbose "Checking path exists: $rootPath$($CallQueue.MusicOnHoldAudioFilePath)"
            }
            catch {
                throw "$errorStringPrefix $CQTypeAccountString $($CallQueue.Name) :: Path error: $rootPath$($CallQueue.MusicOnHoldAudioFilePath) not found/unreachable."
            }
            
            # Path returns true, therefore import and upload the file
            if($PathTest){
                Write-Verbose "Path exists: $rootPath$($CallQueue.MusicOnHoldAudioFilePath)"
                #Generate random filename for audio file
                $MusicOnHoldFilename = "MusicOnHold" + $CallQueue.Name + (Get-Random -Minimum 1 -Maximum 1000) + ".mp3"

                #Lets create duplicate file
                #Build the filename path
                $ParentPath = "$rootPath$($CallQueue.MusicOnHoldAudioFilePath)" | Split-Path -Parent
                $MusicOnHoldCopyFilename = "$ParentPath\$MusicOnHoldFilename"
                Copy-Item -Path "$rootPath$($CallQueue.MusicOnHoldAudioFilePath)" -Destination $MusicOnHoldCopyFilename

                Write-Verbose "$($MusicOnHoldCopyFilename)"

                #Get the audio file from the path
                $MusicOnHoldAudioFileContent = Get-Content -Path $MusicOnHoldCopyFilename -AsByteStream -ReadCount 0
                Write-Verbose "$InfoStringPrefix $CQTypeAccountString $($CallQueue.Name) :: Path exists: MusicOnHold audio file $($MusicOnHoldCopyFilename)"
                Write-Verbose "$InfoStringPrefix $CQTypeAccountString $($CallQueue.Name) :: MusicOnHold audio file $($MusicOnHoldCopyFilename) imported."

                # Import the audio file into Teams
                $MusicOnHoldAudioFile = Import-CsOnlineAudioFile -ApplicationId "HuntGroup” -FileName $MusicOnHoldFilename -Content $MusicOnHoldAudioFileContent

                # Grab the ID - Required for the call queue
                $MusicOnHoldAudioFileID = $MusicOnHoldAudioFile.id

                Write-Verbose "Removing audio file copy: $MusicOnHoldCopyFilename"
                Remove-Item -Path $MusicOnHoldCopyFilename

                Write-Verbose "$InfoStringPrefix $CQTypeAccountString $($CallQueue.Name) :: MusicOnHold audio file uploaded to Teams. ID: $($MusicOnHoldAudioFileID)"
            } else {
                Write-Error "$errorStringPrefix $CQTypeAccountString $($CallQueue.Name) :: MusicOnHold file path $rootPath$($CallQueue.MusicOnHoldAudioFilePath) unreachable."
            }

        }

        # Create call queue
        $NewCQParameters = @{}

        Write-Verbose "Setting Call Queue Name: $($CallQueue.Name)"
        $NewCQParameters.Name = $CallQueue.Name

        # - Ash Ward 2022/02/16 - Need to create this due to the New-CsCallQueue cmdlet requiring the ObjectID and not the UPN
        #Only specify users if it exists in the data
        if($CallQueue.Agents.AgentUPN){
            # Agents
            # Get the GUIDs for each Agent
            $agents = ($x.Agents.split(",") | Get-NASAgentGuid)
            $CallQueue.Agents = $Agents
            Write-Verbose "Setting Agents Guids: $($CallQueue.Agents.AgentGuid)"
            $NewCQParameters.Users = $CallQueue.Agents.AgentGuid
        }else{
            Write-Verbose "Cannot find any agents, skipping"
        }

        Write-Verbose "Setting AgentAlertTime: $($CallQueue.AgentAlertTime)"
        $NewCQParameters.AgentAlertTime = $CallQueue.AgentAlertTime

        if($CallQueue.LanguageID){
            Write-Verbose "Setting LanguageID: $($CallQueue.LanguageID)"
            $NewCQParameters.LanguageID = $CallQueue.LanguageID
        }else{
            Write-Verbose "Setting default en-gb LanguageID due non specified: $($CallQueue.LanguageID)"
            $NewCQParameters.LanguageID = "en-gb"
        }

        if($CallQueue.PresenceBasedRouting -eq "Y"){
            if($CallQueue.RoutingMethod -eq "LongestIdle"){
                #Must set presence based routing to false if longestidle is specified
                Write-Verbose "LongestIdle routing method is specified, setting presence based routing to False"
                Write-Verbose "Setting presence based routing : False"
                $NewCQParameters.RoutingMethod = $CallQueue.RoutingMethod
                $NewCQParameters.PresenceBasedRouting = $False
            }else{
                Write-Verbose "Setting routing method : $($CallQueue.RoutingMethod)"
                Write-Verbose "Setting presence based routing : True"
                $NewCQParameters.RoutingMethod = $CallQueue.RoutingMethod
                $NewCQParameters.PresenceBasedRouting = $True
            }
        }else{
            Write-Verbose "Setting routing method : $($CallQueue.RoutingMethod)"
            Write-Verbose "Setting presence based routing : False"
            $NewCQParameters.RoutingMethod = $CallQueue.RoutingMethod
            $NewCQParameters.PresenceBasedRouting = $False
        }

        if($CallQueue.OverflowThreshold){
            Write-Verbose "Setting OverflowThreshold: $($CallQueue.OverflowThreshold)"
            $NewCQParameters.OverflowThreshold = $CallQueue.OverflowThreshold
        }

        if($CallQueue.TimeoutThreshold){
            Write-Verbose "Setting TimeoutThreshold: $($CallQueue.TimeoutThreshold)"
            $NewCQParameters.TimeoutThreshold = $CallQueue.TimeoutThreshold
        }

        #We must change the string values in the data to a boolean type for Teams

        #Check the data for a N and specify false in the object
        if($CallQueue.AllowOptOut -eq "N"){
            Write-Verbose "Setting Allow Opt Out: $($CallQueue.AllowOptOut)"
            $NewCQParameters.AllowOptOut = $False
        }else{
            Write-Verbose "Setting Allow Opt Out: $($CallQueue.AllowOptOut)"
            $NewCQParameters.AllowOptOut = $True
        }

        #Check the data for a N and specify false in the object
        if($CallQueue.ConferenceMode -eq "N"){
            Write-Verbose "Setting Conference Mode: $($CallQueue.ConferenceMode)"
            $NewCQParameters.ConferenceMode = 0
        }else{
            Write-Verbose "Setting Conference Mode: $($CallQueue.ConferenceMode)"
            $NewCQParameters.ConferenceMode = 1
        }

        #Only specify custom MusicOnHold if it exists in the data
        if((-not [string]::IsNullOrEmpty($CallQueue.MusicOnHoldAudioFilePath)) -and ($CallQueue.UseDefaultMusicOnHold -eq "N")){
            Write-Verbose "Setting custom music on hold"
            Write-verbose "Music on hold file path: $($CallQueue.MusicOnHoldAudioFilePath)"
            Write-Verbose "Setting use default music on hold to: $($CallQueue.UseDefaultMusicOnHold)"
            $NewCQParameters.MusicOnHoldAudioFileId = $MusicOnHoldAudioFileId
            $NewCQParameters.UseDefaultMusicOnHold = $False
        }else{
            Write-Verbose "Setting Default music on hold: True"
            $NewCQParameters.UseDefaultMusicOnHold = $True
        }

        #Only specify custom greeting audio file if it exists in the data
        if($CallQueue.WelcomeMusicAudioFilePath){
            Write-Verbose "Setting welcome music"
            Write-Verbose $CallQueue.WelcomeMusicAudioFilePath
            $NewCQParameters.WelcomeMusicAudioFileId = $WelcomeMusicAudioFileId
        }

        if($CallQueue.OverflowAction -ne "Disconnect"){
            Write-Verbose "OverflowAction not set to Disconnect and therefore setting the new OverflowAction"
            Write-Verbose "Checking if overflow target: $($CallQueue.OverflowActionTarget) is a phone number"
            if($CallQueue.OverflowActionTarget -like "tel:*"){
                Write-Verbose "Overflow target: $($CallQueue.OverflowActionTarget) is a phone number, setting the value"
                $NewCQParameters.OverflowActionTarget = $CallQueue.OverflowActionTarget
                $NewCQParameters.OverflowAction = $CallQueue.OverflowAction
            }else{
                Write-Verbose "Overflow target: $($CallQueue.OverflowActionTarget) is not a phone number"
                if($CallQueue.OverflowActionTarget -like "sip:*"){
                    Write-Verbose "Overflow target: $($CallQueue.OverflowActionTarget) is a sip address"
                    Write-Verbose "Stripping sip: from the overflow target $($CallQueue.OverflowActionTarget) to grab objectID"
                    $OverflowTargetCheck = $CallQueue.OverflowActionTarget.substring(4)
                }else{
                    Write-Verbose "Overflow target is not a sip address"
                    $OverflowTargetCheck = $CallQueue.OverflowActionTarget
                }
                ####
                Write-Verbose "Checking if the overflow target: $($CallQueue.OverflowActionTarget) is valid"
                ####
                if(($OverflowTargetCheck | Get-NASObjectGuid).objguid.guid){
                    $OverflowTargetGuid = ($OverflowTargetCheck | Get-NASObjectGuid).objguid.guid

                    Write-Verbose "Overflow target is valid, setting to: $OverflowTargetGuid"
                    $NewCQParameters.OverflowActionTarget = $OverflowTargetGuid
                    $NewCQParameters.OverflowAction = $CallQueue.OverflowAction
                }else{
                    Write-Error "Overflow target: $OverflowTargetCheck doesn't exist/cannot find GUID."
                }
            }
            #Only specify the OverflowActionTarget if it exists in the data
            if($($CallQueue.OverflowAction) -eq "SharedVoicemail"){
                Write-Verbose "Setting OverflowActionTarget: $($CallQueue.OverflowActionTarget)"
                $NewCQParameters.OverflowActionTarget = $CallQueue.OverflowActionTarget
                Write-Verbose "Setting OverflowSharedVoicemailTranscription: $($CallQueue.EnableOverflowSharedVoicemailTranscription)"
                $NewCQParameters.EnableOverflowSharedVoicemailTranscription = $CallQueue.EnableOverflowSharedVoicemailTranscription
                if($CallQueue.OverflowSharedVoicemailAudioFilePrompt){
                    Write-Verbose "Setting OverflowSharedVoicemailAudioFilePrompt: $($CallQueue.OverflowSharedVoicemailAudioFilePrompt)"
                    $NewCQParameters.OverflowSharedVoicemailAudioFilePrompt = $CallQueue.OverflowSharedVoicemailAudioFilePrompt
                }
    
                if($CallQueue.OverflowSharedVoicemailTextToSpeechPrompt){
                    Write-Verbose "Setting OverflowSharedVoicemailTextToSpeechPrompt: $($CallQueue.OverflowSharedVoicemailTextToSpeechPrompt)"
                    $NewCQParameters.OverflowSharedVoicemailTextToSpeechPrompt = $CallQueue.OverflowSharedVoicemailTextToSpeechPrompt
                }
            }
        }else{
            Write-Verbose "Overflow action is set to: $($CallQueue.OverflowAction), setting value to disconnect"
            $NewCQParameters.OverflowAction = "Disconnect"
        }

        if($CallQueue.TimeoutAction -ne "Disconnect"){
            Write-Verbose "TimeoutAction not set to Disconnect and therefore setting the new TimeoutAction"
            Write-Verbose "Checking if timeout target: $($CallQueue.TimeoutActionTarget) is a phone number"
            if($CallQueue.TimeoutActionTarget -like "tel:*"){
                Write-Verbose "Timeout target: $($CallQueue.TimeoutActionTarget) is a phone number, setting the value"
                $NewCQParameters.TimeoutActionTarget = $CallQueue.TimeoutActionTarget
                $NewCQParameters.TimeoutAction = $CallQueue.TimeoutAction
            }else{
                Write-Verbose "Timeout target: $($CallQueue.TimeoutActionTarget) is not a phone number"
                if($CallQueue.TimeoutActionTarget -like "sip:*"){
                    Write-Verbose "Timeout target: $($CallQueue.TimeoutActionTarget) is a sip address"
                    Write-Verbose "Stripping sip: from the Timeout target $($CallQueue.TimeoutActionTarget) to grab objectID"
                    $TimeoutTargetCheck = $CallQueue.TimeoutActionTarget.substring(4)
                    Write-Verbose "Timeout target set to: $TimeoutTargetCheck"
                }else{
                    Write-Verbose "Timeout target is not a sip address"
                    $TimeoutTargetCheck = $CallQueue.TimeoutActionTarget
                }
                ######
                Write-Verbose "Checking if the Timeout target: $TimeoutTargetCheck is valid"
                ######
                if(($TimeoutTargetCheck | Get-NASObjectGuid).objguid.guid){
                    $TimeoutTargetGuid = ($TimeoutTargetCheck | Get-NASObjectGuid).objguid.guid
                    Write-Verbose "Timeout target is valid, setting to: $TimeoutTargetCheck"
                    $NewCQParameters.TimeoutActionTarget = $TimeoutTargetGuid
                    $NewCQParameters.TimeoutAction = $CallQueue.TimeoutAction
                }else{
                    Write-Error "Timeout target: $TimeoutTargetCheck doesn't exist/cannot find GUID."
                }
            }
            ##Only specify the TimeoutActionTarget if it exists in the data
            #if(($CallQueue.TimeoutAction -eq "Forward" -or "Voicemail" -or "SharedVoicemail") -and ($CallQueue.TimeoutActionTarget)){
            #    Write-Verbose "Setting TimeoutActionTarget: $($CallQueue.TimeoutActionTarget)"
            #    $NewCQParameters.TimeoutActionTarget = $CallQueue.TimeoutActionTarget
            #}

            #Only specify the TimeoutActionTarget if it exists in the data
            if($($CallQueue.TimeoutAction) -eq "SharedVoicemail"){
                Write-Verbose "Setting TimeoutActionTarget: $($CallQueue.TimeoutActionTarget)"
                $NewCQParameters.TimeoutActionTarget = $CallQueue.TimeoutActionTarget
                Write-Verbose "Setting TimeoutSharedVoicemailTranscription: $($CallQueue.EnableTimeoutSharedVoicemailTranscription)"
                $NewCQParameters.EnableTimeoutSharedVoicemailTranscription = $CallQueue.EnableTimeoutSharedVoicemailTranscription
                if($CallQueue.TimeoutSharedVoicemailAudioFilePrompt){
                    Write-Verbose "Setting TimeoutSharedVoicemailAudioFilePrompt: $($CallQueue.TimeoutSharedVoicemailAudioFilePrompt)"
                    $NewCQParameters.TimeoutSharedVoicemailAudioFilePrompt = $CallQueue.TimeoutSharedVoicemailAudioFilePrompt
                }
                if($CallQueue.TimeoutSharedVoicemailTextToSpeechPrompt){
                    Write-Verbose "Setting TimeoutSharedVoicemailTextToSpeechPrompt: $($CallQueue.TimeoutSharedVoicemailTextToSpeechPrompt)"
                    $NewCQParameters.TimeoutSharedVoicemailTextToSpeechPrompt = $CallQueue.TimeoutSharedVoicemailTextToSpeechPrompt
                }
            }
        }else{
            $NewCQParameters.TimeoutAction = "Disconnect"
        }

        #Only specify the DistributionLists if it exists in the data
        if($CallQueue.DistributionLists){
            Write-Verbose "Setting DistributionLists: $($CallQueue.DistributionLists)"
            $NewCQParameters.DistributionLists = $CallQueue.DistributionLists
        }

        #Create the call queue from the above parameters
        $NewCallQueue = New-CsCallQueue @NewCQParameters -ErrorAction Stop

        Write-Verbose "Setting Timeout Target: $($NewCQParameters.TimeoutActionTarget) Setting Overflow Target: $($NewCQParameters.OverflowActionTarget)"

        if($NewCQParameters.TimeoutActionTarget){
            Write-Verbose "Setting timeout action target: $($NewCQParameters.TimeoutActionTarget)"
            Set-CsCallQueue -Identity $NewCallQueue.Identity -TimeoutAction $($NewCQParameters.TimeoutAction) -TimeoutActionTarget "$($NewCQParameters.TimeoutActionTarget)"
        }else{
            Write-Verbose "No need to set timeout action target, no result"
        }
        if($NewCQParameters.OverflowActionTarget){
            Write-Verbose "Setting overflow action target: $($NewCQParameters.OverflowActionTarget)"
            Set-CsCallQueue -Identity $NewCallQueue.Identity -OverflowAction $($NewCQParameters.OverflowAction) -OverflowActionTarget "$($NewCQParameters.OverflowActionTarget)"
        }else{
            Write-Verbose "No need to set overflow action target, no result"
        }
        
        Write-Verbose "$InfoStringPrefix $CQTypeAccountString $($CallQueue.Name) - Call queue has now been created."
        #Pass the call queue ObjectID back to the GUID of the call queue object
        $CallQueue.guid = $NewCallQueue.Identity
    }else{

        #The call queue exists, skip to associating the resource account 
        $NewCallQueue = Get-CsCallQueue -NameFilter "$($CallQueue.Name)"
        Write-Verbose "$InfoStringPrefix $CQTypeAccountString $($CallQueue.Name) - Call queue already exists, skipping to resource account association."
        $CallQueue.guid = $NewCallQueue.Identity
    }

    #Output the call queue object
    $CallQueue
}

function New-NasTeamsAutoAttendant {

    [CmdletBinding()]

    $VerbosePreference = "Continue"
    $OverflowUser = "RA_CQ_OVERFLOW@wardlabnet.onmicrosoft.com"
    $Option1Acc = "TestUser@ward-corp.co.uk"

    #Variables
    $DisplayName = 'ModTestAutoAttendant'
    $TenantDomain = 'wardlabnet.onmicrosoft.com'
    $AAAppID = "ce933385-9390-45d1-9512-c8d228074e07"
    $AAPrefix = "raaa_"

    $errorStringPrefix = "[ERROR]"
    $InfoStringPrefix = "[INFO]"
    $RATypeAccountString = " :: RESOURCE ACCOUNT :: "
    $AATypeAccountString = " :: AUTO ATTENDANT :: "

    if(Get-CsOnlineUser -Identity "$AAPrefix$($DisplayName)@$TenantDomain"){
        Write-Verbose "$InfoStringPrefix $RATypeAccountString $AAPrefix$($DisplayName)@$TenantDomain already exists, skipping to auto attendant creation."
    } else{
        Write-Verbose "$InfoStringPrefix $RATypeAccountString $AAPrefix$($DisplayName)@$TenantDomain doesn't exist, moving to creation."
        $NewAARAParameters = @{
            UserPrincipalName = "$AAPrefix$($DisplayName)@$TenantDomain"
            ApplicationId = "$($AAAppID)"
            DisplayName = "$AAPrefix$($DisplayName)"
        }
        $NewAARA = New-CsOnlineApplicationInstance @NewAARAParameters
        Write-Verbose "$InfoStringPrefix $RATypeAccountString $AAPrefix$($DisplayName)@$TenantDomain now created. Moving to Auto Attendant creation."
    }

    Write-Verbose "$InfoStringPrefix $RATypeAccountString Checking if $AAPrefix$($DisplayName)@$TenantDomain has been created..."

    Do{
        $NewAARA = Get-CsOnlineApplicationInstance -Identity "$AAPrefix$($DisplayName)@$TenantDomain"
    }
    while($null -eq $NewAARA){
        if($Check -le '10'){
            $Check++
            start-sleep -s 10
            $NewAARA = Get-CsOnlineApplicationInstance -Identity "$AAPrefix$($DisplayName)@$TenantDomain"
    }else{
    Throw "$ErrorStringPrefix $RATypeAccountString Resource account not created yet."
    }
    }

    if(Get-CsAutoAttendant -NameFilter "AA_$($DisplayName)"){
        Write-Verbose "$InfoStringPrefix $AATypeAccountString AA_$($DisplayName)@$TenantDomain already exists, skipping..."
    } else{
        # Options
        $aaLanguage = 'en-GB'
        $aaTimezone = 'GMT Standard Time'
        $greetingText = "Welcome to Test Lab"
        $mainMenuText = "To talk to Sales, please press 1. To talk to User2 press 2."
        $afterHoursText = "Sorry Test Lab is closed. Please call back during week days from 9AM to 5PM. Goodbye!"
        $user1Id = (Get-CsOnlineUser -Identity $Option1Acc).Identity
        $salesCQappinstance = (Get-CsOnlineUser -Identity $OverflowUser).Identity # one of the application instances associated to the Call Queue
        $tr1 = New-CsOnlineTimeRange -Start 09:00 -End 17:00
        Write-Verbose "$InfoStringPrefix $AATypeAccountString AA_$($DisplayName)@$TenantDomain - Configured Auto Attendant options"

        # After hours
        $afterHoursSchedule = New-CsOnlineSchedule -Name "After Hours" -WeeklyRecurrentSchedule -MondayHours @($tr1) -TuesdayHours @($tr1) -WednesdayHours @($tr1) -ThursdayHours @($tr1) -FridayHours @($tr1) -Complement
        $afterHoursGreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $afterHoursText
        $afterHoursMenuOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
        $afterHoursMenu = New-CsAutoAttendantMenu -Name "AA menu1" -MenuOptions @($afterHoursMenuOption)
        $afterHoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours" -Menu $afterHoursMenu -Greetings @($afterHoursGreetingPrompt)
        $afterHoursCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type AfterHours -ScheduleId $afterHoursSchedule.Id -CallFlowId $afterHoursCallFlow.Id
        Write-Verbose "$InfoStringPrefix $AATypeAccountString AA_$($DisplayName)@$TenantDomain - Configured Auto Attendant after hours options"

        # Business hours menu options
        $sales = New-CsAutoAttendantCallableEntity -Identity $salesCQappinstance -Type applicationendpoint
        $user1 = New-CsAutoAttendantCallableEntity -Identity $user1Id -Type User
        $menuOption0 = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone0 -CallTarget $sales
        $menuOption1 = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone1 -CallTarget $user1
        Write-Verbose "$InfoStringPrefix $AATypeAccountString AA_$($DisplayName)@$TenantDomain - Configured Auto Attendant menu options"

        # Business hours menu
        $greetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $greetingText
        $menuPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $mainMenuText
        $menu = New-CsAutoAttendantMenu -Name "AA menu2" -Prompts @($menuPrompt) -EnableDialByName -MenuOptions @($menuOption0,$menuOption1)
        Write-Verbose "$InfoStringPrefix $AATypeAccountString AA_$($DisplayName)@$TenantDomain - Configured Auto Attendant menu"

        # And the call flow
        $DefaultCallFlowPArameters = @{
            Name = "$DisplayName Default"
            Menu = $menu
            Greetings = $greetingPrompt
        }
        $defaultCallFlow = New-CsAutoAttendantCallFlow @DefaultCallFlowPArameters
        Write-Verbose "$InfoStringPrefix $AATypeAccountString AA_$($DisplayName)@$TenantDomain - Configured Auto Attendant call flow"

        # You have all the objects
        # Now, you can create an auto attendant
        $AutoAttendantParameters = @{
            Name = "AA_$DisplayName"
            LanguageId = $aaLanguage
            TimeZoneId = $aaTimezone
            DefaultCallFlow = $defaultCallFlow
            CallFlows = @($afterHoursCallFlow)
            CallHandlingAssociations = @($afterHoursCallHandlingAssociation)
            ErrorAction = 'Stop'
        }

        $NewAA = New-CsAutoAttendant @AutoAttendantParameters
        Write-Verbose "$InfoStringPrefix $AATypeAccountString AA_$($DisplayName)@$TenantDomain - Auto Attendant created."
    }

    $NewAA = Get-CsAutoAttendant -NameFilter "AA_$($DisplayName)"

    $AppInstanceID = $NewAARA.ObjectID

    if(Get-CsOnlineApplicationInstanceAssociation -Identity $AppInstanceID -ErrorAction SilentlyContinue){
        Write-Verbose "$InfoStringPrefix $AATypeAccountString AA_$($DisplayName)@$TenantDomain already associated with resource account $AAPrefix$($DisplayName)@$TenantDomain"
    }else{
        # Last but not least, the association
        $AAAssociationParameters = @{
            Identities = @($NewAARA.ObjectId)
            ConfigurationId = $NewAA.Identity
            ConfigurationType = 'AutoAttendant'
            ErrorAction = 'Stop'
        }
        $AppInstanceResults = New-CsOnlineApplicationInstanceAssociation @AAAssociationParameters
        Write-Verbose "$InfoStringPrefix $AATypeAccountString AA_$($DisplayName)@$TenantDomain associated with resource account $AAPrefix$($DisplayName)@$TenantDomain"
    }
}

function Confirm-NasValidTarget{

    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [string]
        $Target
    )

    try{
        [guid]::Parse($Target.substring(4).split("@")[0])
        [bool]$False
    }catch{
        [bool]$True
    }
}

function Import-NasAACQData {

    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$rootFolder,
        [parameter()]
        [switch]$Interactive,
        [parameter()]
        [string]$ffmpeglocation,
        [parameter()]
        [string]$AARAPrefix,
        [parameter()]
        [string]$CQRAPrefix,
        [parameter()]
        [string]$TenantDomain,
        [Parameter()]
        [switch]$SkipAudio
    )
    
    #$rootFolder = "C:\Users\Ash Ward\OneDrive - Modality Systems\Documents\Chris\RgsImportExport"
    Write-Verbose "Importing queues from: $rootFolder\Queues.xml"
    $Queues = Import-Clixml -Path "$rootFolder\Queues.xml" | Sort-Object name
    
    Write-Verbose "Importing agent groups from: $rootFolder\AgentGroups.xml"
    $AgentGroups = Import-Clixml -Path "$rootFolder\AgentGroups.xml" | Sort-Object name

    Write-Verbose "Importing workflows from: $rootFolder\Workflows.xml"
    $Workflows = Import-Clixml -Path "$rootFolder\Workflows.xml" | Sort-Object name

    Write-Verbose "Importing business hours from: $rootFolder\HoursOfBusiness.xml"
    $hours = Import-Clixml -Path "$rootFolder\HoursOfBusiness.xml"

    $ImportWorkflows = $Workflows | ForEach-Object{

        Write-Verbose "Building workflow object: $($_.Name) - $($_.Identity.InstanceId.Guid)"
        $fileLocation = ""

        if($_.CustomMusicOnHoldFile.OriginalFileName){
            Write-Verbose "TRUE: $($_.CustomMusicOnHoldFile.UniqueName)"
            Write-Verbose "Importing custom music on hold file: $($_.CustomMusicOnHoldFile.OriginalFileName)"
            $fileLocation = "\audio\{0}\{1}{2}" -f $_.Identity.InstanceId.Guid, $_.CustomMusicOnHoldFile.UniqueName, $_.CustomMusicOnHoldFile.OriginalFileName.substring($_.CustomMusicOnHoldFile.OriginalFileName.lastindexof("."))
        }else{
            Write-Verbose "No custom music on hold file specified, setting file location to null"
            $fileLocation = ""
        }

        if($fileLocation -like "\*"){
            Write-Verbose "Custom music on hold specified, setting default music on hold to false"
            $UseDefaultMusicOnHold = "N"
        }else{
            Write-Verbose "No custom music on hold specified, setting default music on hold to true"
            $UseDefaultMusicOnHold = "Y"
        }

        if($_.CustomMusicOnHoldFile.UniqueName){
            Write-Verbose "Custom music on hold specified, setting file id"
            $CustomMusicOnHoldFileID = $_.CustomMusicOnHoldFile.UniqueName
        }else{
            Write-Verbose "Custom music on hold not specified, setting to null"
            $CustomMusicOnHoldFileID = ""
        }

        if($_.CustomMusicOnHoldFile.OriginalFileName){
            Write-Verbose "Custom music on hold specified, setting filename"
            $CustomMusicOnHoldFileName = $_.CustomMusicOnHoldFile.OriginalFileName
        }else{
            Write-Verbose "Custom music on hold not specified, setting filename to null"
            $CustomMusicOnHoldFileName = ""
        }

        if($_.NonBusinessHoursAction.Uri){
            if($_.NonBusinessHoursAction.Uri -like "sip:+*"){
                Write-Verbose "Non business hours action target: $($_.NonBusinessHoursAction.Uri) is a phone number, converting to tel:+"
                $e164num1 = $($_.NonBusinessHoursAction.Uri).substring(4).split("@")[0]
                $NonBusinessHoursActionUri = "tel:" + $e164num1
                Write-Verbose "Non business hours action target converted: $NonBusinessHoursActionUri"
            }else{
                Write-Verbose "$($_.NonBusinessHoursAction.Uri) not a phone number, passing value back"
                if((Confirm-NasValidTarget -Target $_.NonBusinessHoursAction.Uri) -eq $True){
                    Write-Verbose "Importing non business hours action target: $($_.NonBusinessHoursAction.Uri)"
                    $NonBusinessHoursActionUri = $_.NonBusinessHoursAction.Uri
                }else{
                    Write-Verbose "Invalid non business hours action target: $($_.NonBusinessHoursAction.Uri)"
                    $NonBusinessHoursActionUri = "Invalid Target"
                }
            }
        }else{
            Write-Verbose "No non business hours action target specified, setting value to null"
            $NonBusinessHoursActionUri = ""
        }

        if($_.DefaultAction.Uri){
            if($_.DefaultAction.Uri -like "sip:+*"){
                Write-Verbose "Default action target: $($_.DefaultAction.Uri) is a phone number, converting to tel:+"
                $e164num2 = $($_.DefaultAction.Uri).substring(4).split("@")[0]
                $DefaultActionTargetUri = "tel:" + $e164num2
                Write-Verbose "Default action target converted: $DefaultActionTargetUri"
            }else{
                Write-Verbose "$($_.DefaultAction.Uri) not a phone number, passing value back"
                if((Confirm-NasValidTarget -Target $_.DefaultAction.Uri) -eq $True){
                    Write-Verbose "Importing default action target: $($_.DefaultAction.Uri)"
                    $DefaultActionTargetUri = $_.DefaultAction.Uri
                }else{
                    Write-Verbose "Invalid default action target: $($_.DefaultAction.Uri)"
                    $DefaultActionTargetUri = "Invalid Target"
                }
            }
        }else{
            Write-Verbose "No default action target specified, setting value to null"
            $DefaultActionTargetUri = ""
        }

        if($_.DefaultAction.QueueID.InstanceId.Guid){
            $DefaultActionQueueID = $_.DefaultAction.QueueID.InstanceId.Guid
        }else{
            Write-Verbose "No default action queue ID specified, setting value to null"
            $DefaultActionQueueID = ""
        }

        if($_.DefaultAction.Question.element){
            $DefaultActionQuestion = $_.DefaultAction.Question.element
        }else{
            Write-Verbose "No default action question specified, setting value to null"
            $DefaultActionQuestion = ""
        }

        if($_.NonBusinessHoursAction.Prompt.AudioFilePrompt){
            $NonBusinessHoursActionAudioFilePrompt = $_.NonBusinessHoursAction.Prompt.AudioFilePrompt
        }else{
            Write-Verbose "No non business hours audio file specified, setting value to null"
            $NonBusinessHoursActionAudioFilePrompt = ""
        }

        if($_.NonBusinessHoursAction.Prompt.TextToSpeechPrompt){
            $NonBusinessHoursActionTTSPrompt = $_.NonBusinessHoursAction.Prompt.TextToSpeechPrompt
        }else{
            Write-Verbose "No non business hours Text to Speech specified, setting value to null"
            $NonBusinessHoursActionTTSPrompt = ""
        }

        if($_.NonBusinessHoursAction.Question){
            $NonBusinessHoursActionQuestion = $_.NonBusinessHoursAction.Question
        }else{
            Write-Verbose "No non business hours question specified, setting value to null"
            $NonBusinessHoursActionQuestion = ""
        }

        if($_.NonBusinessHoursAction.QueueID.InstanceId.Guid){
            $NonBusinessHoursQueueID = $_.NonBusinessHoursAction.QueueID.InstanceId.Guid
        }else{
            Write-Verbose "No non business hours queue ID specified, setting value to null"
            $NonBusinessHoursQueueID = ""
        }

        if($_.LineUri){
            $LineURI = $_.LineUri
        }else{
            Write-Verbose "No line uri specified, setting value to null"
            $LineURI = $null
        }

        $_.Name = $_.Name.replace("_"," ")

        # Let's create the 'cleaned' name, ready for the Teams import
        Write-Verbose "Cleaning the imported name to remove special characters"
        $CleanedWorkflowName = Remove-StringSpecialCharacter -String $_.Name -SpecialCharacterToKeep " "

        if(!($AARAPrefix)){
            #Set the resource account prefix
            $AARAPrefix = "raaa-cc-lll-"
        }

        # Need to clean the display name and but keep spaces, resource accounts need spaces removing
        $ResourceAccountUPN = "{0}{1}@{2}" -f $AARAPrefix, $($CleanedWorkflowName.Replace(" ","")), $($TenantDomain.Replace(" ",""))

        Write-Verbose "Configuring the workflow object: $($_.Name) - $($_.Identity.InstanceId.Guid)"
        [PSCustomObject]@{
            Identity = $_.Identity.InstanceId.Guid
            ResourceAccountUPN = $ResourceAccountUPN
            CleanedName = $CleanedWorkflowName
            Name = $_.Name
            PhoneNumber = $LineURI
            LanguageID = $_.Language
            TimeZone = $_.TimeZone
            #HolidaySet = $_.HolidaySetIDList
            BusinessHoursID = $_.BusinessHoursID.InstanceID
            DefaultAction = if($_.DefaultAction.Action -like "TransferTo*"){"Forward"}else{"Disconnect"}
            DefaultActionTargetQueue = $DefaultActionQueueID
            DefaultActionTargetUri = $DefaultActionTargetUri
            DefaultActionQuestions = $DefaultActionQuestion
            NonBusinessHoursAction = if($_.NonBusinessHoursAction.action -like "TransferTo*"){"Forward"}else{"Disconnect"}
            NonBusinessHoursActionAudioFilePrompt = $NonBusinessHoursActionAudioFilePrompt
            NonBusinessHoursActionTextToSpeechPrompt = $NonBusinessHoursActionTTSPrompt
            NonBusinessHoursActionQuestion = $NonBusinessHoursActionQuestion
            NonBusinessHoursActionQueueID = $NonBusinessHoursQueueID
            NonBusinessHoursActionUri = $NonBusinessHoursActionUri
            HolidayAction = if($_.HolidayAction.Action -like "TransferTo*"){"Forward"}else{"Disconnect"}
            UseDefaultMusicOnHold = $UseDefaultMusicOnHold
            CustomMusicOnHoldFileID = $CustomMusicOnHoldFileID
            CustomMusicOnHoldFileName = $CustomMusicOnHoldFileName
            MusicOnHoldAudioFilePath = $fileLocation.replace(".wav",".mp3")
            Active = $_.Active
        }
        Write-Verbose "Configured the workflow object: $($_.Name) - $($_.Identity.InstanceId.Guid)"
    } # end foreach loop - $ImportWorkflows = $Workflows | foreach{

    $AgentGroupsEdit = $AgentGroups | ForEach-Object{

        #Does the agent alert time value exist?
        if($_.AgentAlertTime){
            if(([int]$_.AgentAlertTime -ge 15) -and ([int]$_.AgentAlertTime -le 180)){
                Write-Verbose "Importing agent alert time: $($_.AgentAlertTime)"
                if(($_.AgentAlertTime % 15) -eq 0){
                    $agentAlertTime = $_.AgentAlertTime
                }else{
                    # Setting agent alert time to next multiple of 15 from the value given
                    Write-Verbose "Setting agent alert time to the next multiple of 15 from $($_.AgentAlertTime)"
                    $agentAlertTime = $_.AgentAlertTime - ($_.AgentAlertTime % 15) + 15
                }
            }
        }else{
            #Set a default agent alert time for Teams
            $AgentAlertTime = "15"
        }

        #Check if the routing method exists
        if($_.RoutingMethod){
            #Need to change routing method to Attendant as Parallel doesn't exist in Teams
            if($_.RoutingMethod -eq "Parallel"){
                $RoutingMethod = "Attendant"
            }else{
                #Is the routing method null? If not, then set the routing method value
                if(-not [string]::IsNullOrEmpty($_.RoutingMethod)){
                    $RoutingMethod = $_.RoutingMethod
                }
                else{
                    #set the routing method to default Attendant
                    $RoutingMethod = "Attendant"
                }
            }
        }

        #Check if value exists
        if($_.ParticipationPolicy){
            if($_.ParticipationPolicy -eq "Informal"){
                #If informal, set to False for Teams usage for AllowOptOut
                $ParticipationPolicy = "False"
            }else{
                $ParticipationPolicy = "True"
            }
        }else{
            #Set default opt out to False if the value doesn't exist
            $ParticipationPolicy = "False"
        }

        #Check if agents exist
        if($_.AgentsByUri){
            $Agents = $_.AgentsByUri.AbsolutePath -join ","
        }else{
            #No agents, set to null
            $Agents = ""
        }

        Write-Verbose "Building the agents object: $($_.identity.instanceid.guid)"
        [PSCustomObject]@{
            Identity = $_.identity.instanceid.guid
            Name = $_.Name
            AllowOptOut = $ParticipationPolicy
            AgentAlertTime = $agentAlertTime
            RoutingMethod = $RoutingMethod
            Agents = $Agents
        }
        Write-Verbose "Agents object built: $($_.identity.instanceid.guid)"

    } # - end foreach loop - $AgentGroupsEdit = $AgentGroups | foreach{

    $ImportQueues = foreach($x in $Queues){

        if($x.TimeoutAction.Action -ne "Terminate"){
            Write-Verbose "Timeout action not set to terminate, setting the timeout threshold"
            if(([int]$x.TimeoutThreshold -ge 15) -and ([int]$x.TimeoutThreshold -le 180)){
                Write-Verbose "Timeout threshold is: $($x.TimeoutThreshold) which is > 15 and < 180, checking if the value is a multiple of 15"
                if(($x.TimeoutThreshold % 15) -eq 0){
                    Write-Verbose "Timeout threshold is a multiple of 15, setting the value"
                    $TimeoutThreshold = $x.TimeoutThreshold
                }else{
                    # Setting timeout threshold to next multiple of 15 from the value given
                    Write-Verbose "Timeout threshold is: $($x.TimeoutThreshold) which is NOT a multiple of 15, setting value to next multiple of 15"
                    $TimeoutThreshold = $x.TimeoutThreshold - ($x.TimeoutThreshold % 15) + 15
                }
            }else{
                Write-Verbose "Timeout threshold is: $($x.TimeoutThreshold) which is not > 15 and < 180, checking the value"
                if([int]$x.TimeoutThreshold -lt 15){
                    Write-Verbose "Timeout threshold is: $($x.TimeoutThreshold) which is < 15, setting value to 15"
                    $TimeoutThreshold = "15"
                }
                if([int]$x.TimeoutThreshold -gt 180){
                    Write-Verbose "Timeout threshold is: $($x.TimeoutThreshold) which is > 180, setting value to 180"
                    $TimeoutThreshold = "180"
                }
            }
        }else{
            Write-Verbose "Timeout action is set to terminate, setting the threshold to null"
            $TimeoutThreshold = ""
        }

        if($x.OverflowAction.Action -ne "Terminate"){
            Write-Verbose "Overflow action not set to terminate, setting the overflow threshold"
            if($x.OverflowThreshold -ge 1){
                Write-Verbose "Overflow threshold $($x.OverflowThreshold) is greater than or equal to 1, setting the value"
                $OverflowThreshold = $x.OverflowThreshold
            }else{
                Write-Verbose "Overflow threshold $($x.OverflowThreshold) is not greater than or equal to 1, setting the default value to 1"
                $OverflowThreshold = "1"
            }
        }else{
            Write-Verbose "Overflow action is set to terminate, setting the threshold to null"
            $OverflowThreshold = ""
        }

        if($x.OverflowAction.Uri){
            if($x.OverflowAction.Uri -like "sip:+*"){
                Write-Verbose "Overflow action target: $($x.OverflowAction.Uri) is a phone number, converting to tel:+"
                $e164num3 = $($x.OverflowAction.Uri).substring(4).split("@")[0]
                $OverflowActionUri = "tel:$e164num3"
                Write-Verbose "Converted overflow action target to: $OverflowActionUri"

            }else{
                Write-Verbose "$($x.DefaultAction.Uri) not a phone number, checking if it's valid, if so, passing back"
                if($x.OverflowAction.Uri | Confirm-NasValidTarget){
                    Write-Verbose "Importing overflow action target: $($x.OverflowAction.Uri)"
                    $OverflowActionUri = $x.OverflowAction.Uri
                }else{
                    Write-Verbose "Invalid overflow action target: $($x.OverflowAction.Uri)"
                    $OverflowActionUri = "Invalid Target"
                }
            }
        }else{
            Write-Verbose "No overflow action target specified, setting value to null"
            $OverflowActionUri = ""
        }

        if($x.TimeoutAction.Uri){
            if($x.TimeoutAction.Uri -like "sip:+*"){
                Write-Verbose "Timeout action target: $($x.TimeoutAction.Uri) is a phone number, converting to tel:+"
                $e164num4 = $($x.TimeoutAction.Uri).substring(4).split("@")[0]
                $TimeoutActionUri = "tel:$e164num4"
                Write-Verbose "Converted timeout action target to: $TimeoutActionUri"
            }else{
                Write-Verbose "$($x.TimeoutAction.Uri) not a phone number, checking if it's valid, if so, passing back"
                if($x.TimeoutAction.Uri | Confirm-NasValidTarget){
                    Write-Verbose "Importing timeout action target: $($x.TimeoutAction.Uri)"
                    $TimeoutActionUri = $x.TimeoutAction.Uri
                }else{
                    Write-Verbose "Invalid timeout action target: $($x.TimeoutAction.Uri)"
                    $TimeoutActionUri = "Invalid Target"
                }
            }
        }else{
            Write-Verbose "No timeout action target specified, setting value to null"
            $TimeoutActionUri = ""
        }
        
        $filterAgent = $AgentGroupsEdit.where({$_.identity -eq $x.agentgroupidlist.instanceid.guid})

        #Null the vars
        $UseDefaultMusicOnHoldQueue,$CustomMusicOnHoldOriginalFileNameQueue,$CustomMusicOnHoldFilePathQueue = $null

        #Run if -Interactive is specified to let the user know what to do
        if($Interactive){
            Write-Information "You must choose a queue configuration in the pop up window, otherwise to exit, press cancel twice."
        }
        
        # Set the music variables
        if(($ImportWorkflows.where({$_.DefaultActionTargetQueue -eq $x.Identity.InstanceId.Guid }).UseDefaultMusicOnHold).count -gt 1){
            #If not -Interactive, throw to let the user know there is problem with the data
            if(!($interactive)){
                Write-Error "Call queue: ID: $($x.Identity.InstanceId.Guid) - Name: $($x.Name) found multiple workflows with differing music on hold settings"
                throw "Re-run the cmdlet as ""Import-NasAACQData -Interactive"" to choose which music on hold value."
            }else{
                $returnObject = $null
                #Check the value is present, if it isn't after 2 attempts, throw an error
                while(!($returnObject) -and ($ignore -le 1)){
                    $ignore++
                    $returnObject = ($ImportWorkflows.where({$_.DefaultActionTargetQueue -eq $x.Identity.InstanceId.Guid }) | Out-GridView -OutputMode Single -Title "Choose the queue music configuration required")
                    Write-verbose "UseDefaultMusicOnHold: chosen option $($returnObject.UseDefaultMusicOnHold)"
                }
                if(!($returnObject)){
                    throw "You must choose an option to continue."
                }
                $UseDefaultMusicOnHoldQueue = $returnObject.UseDefaultMusicOnHold
                $CustomMusicOnHoldOriginalFileNameQueue = $returnObject.CustomMusicOnHoldFileName
                $CustomMusicOnHoldFilePathQueue = $returnObject.MusicOnHoldAudioFilePath

                Write-Verbose "Setting UseDefaultMusicOnHold to: $UseDefaultMusicOnHoldQueue"
                if($CustomMusicOnHoldOriginalFileNameQueue -or $CustomMusicOnHoldFilePathQueue){
                    Write-Verbose "Setting MusicOnHoldAudioFilePath to: $CustomMusicOnHoldOriginalFileNameQueue"
                    Write-Verbose "Setting CustomMusicOnHoldOriginalFileName to: $CustomMusicOnHoldFilePathQueue"
                }
            }
        }else{
            if(($ImportWorkflows.where({$_.DefaultActionTargetQueue -eq $x.Identity.InstanceId.Guid }).UseDefaultMusicOnHold) -eq "Y"){
                $UseDefaultMusicOnHoldQueue = "Y"
            }
            if(($ImportWorkflows.where({$_.DefaultActionTargetQueue -eq $x.Identity.InstanceId.Guid }).UseDefaultMusicOnHold) -eq "N"){
                # Set the music on hold values
                Write-Verbose "Setting music on hold values."
                Write-Verbose "`$UseDefaultMusicOnHoldQueue = $($ImportWorkflows.where({$_.DefaultActionTargetQueue -eq $x.Identity.InstanceId.Guid }).UseDefaultMusicOnHold)"
                Write-Verbose "`$CustomMusicOnHoldOriginalFileNameQueue = $($ImportWorkflows.where({$_.DefaultActionTargetQueue -eq $x.Identity.InstanceId.Guid }).CustomMusicOnHoldFileName)"
                Write-Verbose "`$CustomMusicOnHoldFilePathQueue = $($ImportWorkflows.where({$_.DefaultActionTargetQueue -eq $x.Identity.InstanceId.Guid }).MusicOnHoldAudioFilePath)"
                $UseDefaultMusicOnHoldQueue = "N"
                Write-Verbose "Use default music on hold is set to N, setting custom music on hold"
                $CustomMusicOnHoldOriginalFileNameQueue = $ImportWorkflows.where({$_.DefaultActionTargetQueue -eq $x.Identity.InstanceId.Guid }).CustomMusicOnHoldFileName
                $CustomMusicOnHoldFilePathQueue = $ImportWorkflows.where({$_.DefaultActionTargetQueue -eq $x.Identity.InstanceId.Guid }).MusicOnHoldAudioFilePath
            }
            else{
                $UseDefaultMusicOnHoldQueue = "Y"
            }
        }

        if($x.OverflowAction.Action -like "TransferTo*"){
            Write-Verbose "Overflow action: $($x.OverflowAction.Action) set to TransferTo, converting to Forward for Teams usage"
            $OverflowAction = "Forward"
        }else{
            Write-Verbose "Overflow action: $($x.OverflowAction.Action) set to Terminate, converting to Disconnect for Teams usage"
            $OverflowAction = "Disconnect"
        }

        if($x.TimeoutAction.Action -like "TransferTo*"){
            Write-Verbose "Timeout action: $($x.TimeoutAction.Action) set to TransferTo, converting to Forward for Teams usage"
            $TimeoutAction = "Forward"
        }else{
            Write-Verbose "Timeout action: $($x.TimeoutAction.Action) set to Terminate, converting to Disconnect for Teams usage"
            $TimeoutAction = "Disconnect"
        }

        #If there is no Routing Method specified, set to Attendant
        if(!($filterAgent.RoutingMethod)){
            $RoutingMethod = "Attendant"
        }else{
            $RoutingMethod = $filterAgent.RoutingMethod
        }

        #If there is no Agent Alert Time specified, set to default 15
        if(!($filterAgent.AgentAlertTime)){
            $AgentAlertTime = "15"
        }else{
            $AgentAlertTime = $filterAgent.AgentAlertTime
        }

        #If there is no AllowOptOut specified, set to default False
        if(!($filterAgent.AllowOptOut)){
            $AllowOptOut = "N"
        }else{
            if($filterAgent.AllowOptOut -eq $False){
                $AllowOptOut = "N"
            }else{
                $AllowOptOut = "Y"
            }
        }

        Write-Verbose "Building the queue object: $($x.Name) - $($x.Identity.InstanceId.Guid)"

        Write-Verbose "Overflow action target? $OverflowActionUri"
        Write-Verbose "TImeout action target? $TimeoutActionUri"

        # Let's create the 'cleaned' name, ready for the Teams import
        Write-Verbose "Cleaning the imported name to remove special characters"

        $x.name = $x.name.replace("_"," ")

        $CleanedQueueName = Remove-StringSpecialCharacter -String $x.Name -SpecialCharacterToKeep " "

        #If prefix is specified for bespoke requirements
        if(!($CQRAPrefix)){
            #Set the resource account prefix
            $CQRAPrefix = "racq-cc-lll-"
        }

        #Build the resource account upn for bespoke requirements
        $ResourceAccountUPN = "{0}{1}@{2}" -f $CQRAPrefix, $($CleanedQueueName.Replace(" ","")), $($TenantDomain.Replace(" ",""))

        $returncqobj = [PSCustomObject][ordered]@{
            QueueID = $x.Identity.InstanceId.Guid
            ResourceAccountUPN = $ResourceAccountUPN
            CleanedName = $CleanedQueueName
            Name = $x.Name
            RoutingMethod = $RoutingMethod
            PresenceBasedRouting = "N"
            ConferenceMode = "Y"
            AgentAlertTime = $AgentAlertTime
            AllowOptOut = $AllowOptOut
            UseDefaultMusicOnHold = $UseDefaultMusicOnHoldQueue
            CustomMusicOnHoldOriginalFileName = $CustomMusicOnHoldOriginalFileNameQueue
            MusicOnHoldAudioFilePath = $CustomMusicOnHoldFilePathQueue
            Agents = $filterAgent.Agents
            OverflowAction = $OverflowAction
            OverflowThreshold = $OverflowThreshold
            OverflowActionTarget = $OverflowActionUri
            TimeoutAction = $TimeoutAction
            TimeoutThreshold = $TimeoutThreshold
            TimeoutActionTarget = $TimeoutActionUri
        }
        $returncqobj
        Write-Verbose "Built the queue object: $($x.Name) -  $($x.Identity.InstanceId.Guid)"
        
    } # end foreach loop - $ImportQueues = foreach($x in $Queues){

    # Export to excel
    Write-Verbose "Exporting workflows to Excel"
    $ImportWorkflows | Sort-Object name | Export-Excel -Path "$rootFolder\AACQDataImport.xlsx" -WorksheetName "Auto Attendants" -BoldTopRow -AutoSize

    Write-Verbose "Exporting queues to Excel"
    $ImportQueues | Sort-Object name | Export-Excel -Path "$rootFolder\AACQDataImport.xlsx" -WorksheetName "Call Queues" -BoldTopRow -AutoSize

    Write-Verbose "Exporting business hours to Excel"
    $hours | Select-Object @{l="BusinessHoursID";exp={$_.Identity.InstanceID.Guid}}, @{l="Name";exp={$_.Name.split("_")[0]}},
        @{l="MonOpen";e={$_.MondayHours1.OpenTime.ToString()}},@{l="MonClose";e={$_.MondayHours1.CloseTime.ToString()}},
        @{l="TueOpen";e={$_.TuesdayHours1.OpenTime.ToString()}},@{l="TueClose";e={$_.TuesdayHours1.CloseTime.ToString()}},
        @{l="WedsOpen";e={$_.WednesdayHours1.OpenTime.ToString()}},@{l="WedsClose";e={$_.WednesdayHours1.CloseTime.ToString()}},
        @{l="ThursOpen";e={$_.ThursdayHours1.OpenTime.ToString()}},@{l="ThursClose";e={$_.ThursdayHours1.CloseTime.ToString()}},
        @{l="FriOpen";e={$_.FridayHours1.OpenTime.ToString()}},@{l="FriClose";e={$_.FridayHours1.CloseTime.ToString()}},
        @{l="SatOpen";e={$_.SaturdayHours1.OpenTime.ToString()}},@{l="SatClose";e={$_.SaturdayHours1.CloseTime.ToString()}},
        @{l="SunOpen";e={$_.SundayHours1.OpenTime.ToString()}},@{l="SunClose";e={$_.SundayHours1.CloseTime.ToString()}} |
        Export-Excel -Path "$rootFolder\AACQDataImport.xlsx" -WorksheetName "Business Hours" -NoNumberConversion "Name" -BoldTopRow -AutoSize
    
    #Export the languages
    Write-Verbose "Exporting the languages"
    Get-NASTeamsLanguages -rootFolder $rootFolder

    Write-Verbose "Exports complete, location: $rootFolder\AACQDataImport.xlsx"

    if(!($SkipAudio)){
        #Convert and export music files
        Convert-NasImportMusicFile -rootfolder $rootfolder -ffmpegLocation $ffmpeglocation
    }else{
        Write-Verbose "-SkipAudio set, skipping audio conversion."
    }

}

function Convert-NasImportMusicFile{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [string]$rootfolder,
        [string]$ScriptLocation = (Get-Location).Path,
        [string]$ffmpegLocation
    )

    #Grab the workflows that don't default music on hold set to Y
    #Small quirk, some workflows may have 2 queues associated, and therefore two music on hold values, only grab the ones less than 1
    $NewMusicFolders = $ImportWorkflows.where({$_.UseDefaultMusicOnHold -ne "Y" -and $($_.UseDefaultMusicOnHold).count -le 1})

    # Lets loop through the NewMusicFolders, create the queue folders, convert the files and move them into the correct folder
    ForEach($musicid in $NewMusicFolders){

        #Clear the last objects
        $OriginalFile,$OriginalFileString,$OriginalPath,$newFile,$newFolderPath = $null

        # Check if the queue id music folder exists, if not, create it
        if(!(Test-Path -Path $rootfolder\audio\$($musicid.identity))){
            Write-Verbose "Creating audio folder: $rootfolder\audio\$($musicid.identity)"
            $newFolderPath = "$rootfolder\audio\$($musicid.identity)"
            New-Item -Path $newFolderPath -ItemType Directory
        }else{
            Write-Verbose "Folder $rootfolder\audio\$($musicid.identity) already exists"
        }
        # Check if it exists, if it exists then start to build out the new converted files
        if(Test-Path -Path $rootfolder\audio\$($musicid.identity)){
            Write-Verbose "Audio folder already exists, check the location: $rootfolder\audio\$($musicid.identity)"
            Write-Verbose "Root folder: $rootfolder"
            $OriginalPath = "$rootfolder\RGS\Instances\$($musicid.identity)"

            Write-Verbose "Original Path: $OriginalPath"
            $OriginalFile = Get-ChildItem -Path $OriginalPath -Recurse -Filter "$($musicid.CustomMusicOnHoldFileID).wav"

            Write-Verbose "Original File: $OriginalFile"
            # This will be the new file name
            $newFile = "$($OriginalFile.basename).mp3"
            Write-Verbose "Filename $newfile"

            #Change file to string
            $OriginalFileString = $OriginalFile.tostring()

            # Build the new file path
            #$OriginalCompletePath = "$OriginalPath\$OriginalFileString"
            Write-Verbose "Checking new file path: $OriginalFileString"
            
            # Execute ffmpeg to convert the file
            & $ffmpegLocation -i $OriginalFileString $newFile
            Write-Verbose "Original file name: $OriginalFileString"
            Write-Verbose "New file name: $newfile"

            # Specify the destination of the converted file
            $dest = "$rootfolder\audio\$($musicid.identity)"
            $pathofconvertedmusic = "$ScriptLocation\$newFile"
            
            # Move the converted file to the destination $dest
            Move-Item -Path $pathofconvertedmusic -Destination $dest
            Write-Verbose "File $OriginalFileString converted to mp3 - result: $newfile"
        }else{
            Write-Verbose "Folder: $rootfolder\$($musicid.identity) doesn't exist"
        }
    }
    Write-Verbose "All music files converted to mp3"
}