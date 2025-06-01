
$NotificationFunctionsDefs = {

    <#
    .SYNOPSIS
    	Title   : Function Confirm-WinRuntimeTypeLoaded
    	Author  : Jon Damvi
    	Version : 1.0.0
    	Date    : 01.06.2025
    	License : MIT (LICENSE)
    
    	Release Notes: 
    		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).
    
    .DESCRIPTION
    	Checks if specified Windows Runtime Type is loaded in current session.
        Returns boolean True if Windows Runtime Type is loaded. Otherwise returns False.
    
    .PARAMETER TypeName
    	(Mandatory) Specifies Windows Runtime Type Name to check.
    	Expected type: [String]
    
    .INPUTS
    	None
    
    .OUTPUTS
    	[bool] - returns True or False if Windows Runtime Type is loaded.
    
    .EXAMPLE
    	# Usage:
    	PS > Confirm-WinRuntimeTypeLoaded -TypeName 'Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime'
    
    .LINK
    	GitHub Repository: https://github.com/...
    #>
    Function Confirm-WinRuntimeTypeLoaded {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [ValidatePattern('^[A-Z][A-Z0-9.]*?[A-Z0-9]\s*?,\s*?[A-Z][A-Z0-9.]*?[A-Z0-9]\s*?,\s*?ContentType=WindowsRuntime$')]
            [string]$TypeName
        )
        Try {
            $Type = [Type]::GetType($TypeName, $false)
            If ($null -ne $Type) {
                Write-Verbose "Type '$TypeName' is already loaded."
                return $true
            } Else {
                return $false
            }
        }
        Catch {
            return $false
        }
    }

    <#
    .SYNOPSIS
    	Title   : Function Load-WinRuntimeType
    	Author  : Jon Damvi
    	Version : 1.0.0
    	Date    : 01.06.2025
    	License : MIT (LICENSE)
    
    	Release Notes: 
    		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).
    
    .DESCRIPTION
    	Loads specified Windows Runtime Type in current session.
    
    .PARAMETER TypeName
    	(Mandatory) Specifies Windows Runtime Type Name to load.
    	Expected type: [String]
    
    .INPUTS
    	None
    
    .OUTPUTS
    	None
    
    .EXAMPLE
    	# Usage:
    	PS > Load-WinRuntimeType -TypeName 'Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime'
    
    .LINK
    	GitHub Repository: https://github.com/...
    #>
    Function Load-WinRuntimeType {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [ValidatePattern('^[A-Z][A-Z0-9.]*?[A-Z0-9]\s*?,\s*?[A-Z][A-Z0-9.]*?[A-Z0-9]\s*?,\s*?ContentType=WindowsRuntime$')]
            [string]$TypeName
        )
        If (Confirm-WinRuntimeTypeLoaded $TypeName) { return }
        Try { # Cast to the type to force loading
            [void][Type]::GetType($typeName, $true)
            Write-Output "Loaded '$TypeName' Type Successfully!"
        }
        Catch {
            $ErrorMessage = "$($_.Exception.Message)`n$($_ | Format-List | Out-String)`nError Trace:`n$($_.ScriptStackTrace)"
            Write-Error "Failed to load WinRT type '$TypeName': $ErrorMessage"
            Throw $_
        }
    }

    <#
    .SYNOPSIS
    	Title   : Function Confirm-WavFileType
    	Author  : Jon Damvi
    	Version : 1.0.0
    	Date    : 01.06.2025
    	License : MIT (LICENSE)
    
    	Release Notes: 
    		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).
    
    .DESCRIPTION
    	Checks specified file signature if it is valid "wav" file.
    
    .PARAMETER FilePath
    	(Mandatory) Specifies file to check file signature for.
    	Expected type: [String]
    
    .INPUTS
    	None
    
    .OUTPUTS
    	[bool] - returns True if file is valid "wav" file, otherwise returns False.
    
    .EXAMPLE
    	# Usage Case ... Example Description :
    	PS > Confirm-WavFileType -FilePath "C:\Windows\Media\Ring05.wav"
    
    .LINK
    	GitHub Repository: https://github.com/...
    #>
    Function Confirm-WavFileType {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$FilePath
        )
        If (-Not (Test-Path $FilePath)) {
            return $false
        }
        $FS = [System.IO.File]::OpenRead($FilePath)
        $Buffer = New-Object byte[] 12
        $FS.Read($Buffer, 0, 12) | Out-Null
        $FS.Close()

        [string]$RIFF = [System.Text.Encoding]::ASCII.GetString($Buffer, 0, 4)
        [string]$WAVE = [System.Text.Encoding]::ASCII.GetString($Buffer, 8, 4)

        If ($RIFF -eq "RIFF" -and $WAVE -eq "WAVE") {
            return $true
        } Else {
            return $false
        }
    }

    <#
    .SYNOPSIS
    	Title   : Function Confirm-PngFileType
    	Author  : Jon Damvi
    	Version : 1.0.0
    	Date    : 01.06.2025
    	License : MIT (LICENSE)
    
    	Release Notes: 
    		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).
    
    .DESCRIPTION
    	Checks specified file signature if it is valid "png" file.
    
    .PARAMETER FilePath
    	(Mandatory) Specifies file to check file signature for.
    	Expected type: [String]
    
    .INPUTS
    	None
    
    .OUTPUTS
    	[bool] - returns True if file is valid "png" file, otherwise returns False.
    
    .EXAMPLE
    	# Usage Case ... Example Description :
    	PS > Confirm-PngFileType -FilePath "C:\Windows\ImmersiveControlPanel\images\Devices.png"
    
    .LINK
    	GitHub Repository: https://github.com/...
    #>
    Function Confirm-PngFileType {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$FilePath
        )
        If (-Not (Test-Path $FilePath)) {
            return $false
        }
        #$Buffer = Get-Content -Path $FilePath -Encoding Byte -TotalCount 8 -Force

        $FS = [System.IO.File]::OpenRead($FilePath)
        $Buffer = New-Object byte[] 8
        $FS.Read($Buffer, 0, 8) | Out-Null
        $FS.Close()

        $PngSignature = [byte[]](0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A)
        $Equal = [System.Linq.Enumerable]::SequenceEqual($Buffer, $PngSignature)

        If ($Equal) {
            return $true
        } Else {
            return $false
        }
    }

    <#
    .SYNOPSIS
    	Title   : Function Show-Notification
    	Author  : Jon Damvi
    	Version : 1.0.0
    	Date    : 01.06.2025
    	License : MIT (LICENSE)
    
    	Release Notes: 
    		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).
    
    .DESCRIPTION
    	Shows Windows Toast Notification with specified custom text message, title, icon.
        Notification Types: 'default' - will pop-up for short time up to 5 seconds and close,
        'reminder' - will pop-up for longer time up to 20 seconds and close, 'alarm' - works
        as 'default' type, 'incomingCall' - will pop-up and be shown until user dismisses manually,
        'urgent' - on first run will trigger system notification to allow receiving urgent
        notification from this app - specified by SourceName parameter. 'urgent' notification has
        red exclamation mark. Scheduling notifications is currently not implemented.

        Toast notification style is configured dynamically based on supplied arguments.
        If path to valid ".png" file specified by IconPath parameter, then it will be shown in
        toast notification. If path to valid ".wav" file specified by AudioPath parameter, then
        it will be played in toast notification. Icon will not be shown if IconPath path is
        invalid or invalid file type specified. Audio will be muted if AudioPath path is invalid
        or invalid file type specified. If AudioPath not specified, then default sound will be
        played based on notificaiton type.
    
    .PARAMETER Message
    	(Mandatory) Specifies Toast Notification Message to show.
    	Expected type: [String]
    
    .PARAMETER SourceName
    	(Mandatory) Specifies Toast Notification Message Source (Application Name).
    	Expected type: [String]
    
    .PARAMETER Title
    	(Optional) Specifies Toast Notification Title.
    	Expected type: [String]
    
    .PARAMETER Muted
    	(Optional) Disables Toast Notification Sound.
    	Expected type: [switch]
    
    .PARAMETER IconPath
    	(Optional) Specifies Toast Notification Icon.
    	Expected type: [String]
    
    .PARAMETER AudioPath
    	(Optional) Specifies Toast Notification Audio to Play (WAV file).
    	Expected type: [String]
    
    .PARAMETER Type
    	(Optional) Specifies Toast Notification Type.
    	Expected type: [String]
    	Allowed Values: default, reminder, alarm, incomingCall, urgent
    	Default Value: 'default'
    
    .PARAMETER ExpirationTime
    	(Optional) Currently not supported. Specifies Toast Notification Expiration Time. 
        If Toast Notification is scheduled, then expiration time should expire toast notification
        if occurs before schedule time and thus prevent notification from showing.
    	Expected type: [DateTimeOffset]
    
    .INPUTS
    	None
    
    .OUTPUTS
    	None
    
    .EXAMPLE
    	# Quick Toast Notification - default sound, no icon.
    	Show-Notification -Message "Notification Message" -SourceName "MyApp"

    .EXAMPLE
    	# Quick Toast Notification - no sound, no icon.
    	Show-Notification -Message "Notification Message" -SourceName "MyApp" -Muted

    .EXAMPLE
    	# Long Toast Notification - default sound, no icon.
    	Show-Notification -Message "Notification Message" -SourceName "MyApp" -Type "reminder"

    .EXAMPLE
    	# Persistent Toast Notification - default sound, no icon.
    	Show-Notification -Message "Notification Message" -SourceName "MyApp" -Type "incomingCall"

    .EXAMPLE
    	# Persistent Toast Notification - custom title, default sound, no icon.
    	Show-Notification -Message "Notification Message" -SourceName "MyApp" -Title "Reminder:" -Type "incomingCall"

    .EXAMPLE 
    	# Persistent Custom Toast Notification - custom title, custom sound, custom icon.
        $NotificationParams = @{
        SourceName = "MyApp"
        Title      = "Reminder:"
        Message    = "Notification Message"
        IconPath   = "C:\Windows\System32\wpcatltoast.png"
        AudioPath  = "C:\Windows\Media\Alarm07.wav"
        Type = 'incomingCall'
    }
    Show-Notification @NotificationParams
    
    .LINK
    	GitHub Repository: https://github.com/...
    #>
    Function Show-Notification {
        [CmdletBinding()]
        Param (
            [Parameter(Mandatory)]
            [string]$Message,

            [Parameter(Mandatory)]
            [string]$SourceName,

            [Parameter()]
            [string]$Title,

            [Parameter()]
            [switch]$Muted,

            [Parameter()]
            [string]$IconPath,

            [Parameter()]
            [string]$AudioPath,

            [Parameter()]
            [ValidateSet('default','reminder','alarm','incomingCall','urgent')]
            [string]$Type = 'default',

            [Parameter()]
            [ValidateNotNullOrEmpty()]
            [DateTimeOffset]$ExpirationTime
        )
        Try {
            # Load required 'Windows.UI.Notifications assembly' and 'Windows.Data.Xml.Dom.XmlDocument' Windows Runtime Assemblies
            Load-WinRuntimeType 'Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime'
            Load-WinRuntimeType 'Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType=WindowsRuntime'

            $template = $null

            $AudioPathValid = ($AudioPath -And (Test-Path $AudioPath) -And (Confirm-WavFileType $AudioPath))
            $IconPathValid  = ($IconPath  -And (Test-Path $IconPath)  -And (Confirm-PngFileType $IconPath))

            If ($IconPathValid) {
                If ($Title) {
                    $template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastImageAndText02)
                } Else {
                    $template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastImageAndText01)
                }
            } Else {
                If ($Title) {
                    $template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
                } Else {
                    $template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText01)
                }
            }

            $xml = $template.GetXml()
            $ToastXml = [xml]$xml

            $textElement = $ToastXml.toast.visual.binding.text

            If ($Type -ne 'Default') {
                $ToastElement = $ToastXml.toast
                $ToastElement.SetAttribute("scenario", $Type)
                If ($Type -eq 'reminder') {
                    $ToastElement.SetAttribute("duration", 'long')
                }
            }

            If ($Title) {
                ($textElement | ? { $_.id -eq "1" }).AppendChild($ToastXml.CreateTextNode($Title))   | Out-Null
                ($textElement | ? { $_.id -eq "2" }).AppendChild($ToastXml.CreateTextNode($Message)) | Out-Null
            } Else {
                ($textElement | ? { $_.id -eq "1" }).AppendChild($ToastXml.CreateTextNode($Message)) | Out-Null
            }

            If ($IconPathValid) {
                $imageElement = $ToastXml.toast.visual.binding.image
                $imageElement.SetAttribute("src", "file:///$IconPath")
            }

            $audioElement = $ToastXml.CreateElement("audio")
            If (-Not $Muted -And -Not $AudioPath) {
                # Default Toast Notification Audio will be played
            } ElseIf ($Muted -OR ($AudioPath -And -Not $AudioPathValid)) {
                $audioElement.SetAttribute("silent", "true")    
            } ElseIf ((-Not $Muted) -And $AudioPath -And $AudioPathValid) {
                $audioElement.SetAttribute("src", "file:///$AudioPath")
            }

            $ToastXml.toast.AppendChild($audioElement) | Out-Null

            # Create toast notification and show it
            $xmlDoc = New-Object Windows.Data.Xml.Dom.XmlDocument
            $xmlDoc.LoadXml($ToastXml.OuterXml)
            $Toast = [Windows.UI.Notifications.ToastNotification]::new($xmlDoc)

            If ($ExpirationTime) {
                $Toast.ExpirationTime = $ExpirationTime
            }

            $Toast.Tag = "PowerShell"
            $Toast.Group = "PowerShell"

            $Notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($SourceName)
            $Notifier.Show($Toast)
            
            Write-Output "Toast notification displayed successfully!"
        }
        Catch {
            $ErrorMessage = "$($_.Exception.Message)`n$($_ | Format-List | Out-String)`nError Trace:`n$($_.ScriptStackTrace)"
            Write-Error "Failed to show notification: $ErrorMessage"
            Throw $_
        }
    }
}

$RunspaceFunctionsDefs = {
    <#
    .SYNOPSIS
    	Title   : Function New-Runspace
    	Author  : Jon Damvi
    	Version : 1.0.0
    	Date    : 01.06.2025
    	License : MIT (LICENSE)
    
    	Release Notes: 
    		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).
    
    .DESCRIPTION
    	Creates new Runspace.
    
    .PARAMETER Interactive
    	(Optional) Enables current user desktop interaction for Runspace.
    	Expected type: [switch]
        Default Value: $false
    
    .INPUTS
    	None
    
    .OUTPUTS
    	System.Management.Automation.Runspaces.LocalRunspace
    
    .EXAMPLE
        # Create new interactive Runspace:
    	$Runspace = New-Runspace -Interactive $Interactive
    
    .LINK
    	GitHub Repository: https://github.com/...
    #>
    Function New-Runspace {
        [CmdletBinding()]
        Param(
            [Parameter()]
            [switch]$Interactive = $false
        )
        $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        $Runspace = [runspacefactory]::CreateRunspace($initialSessionState)
        If ($Interactive) {
            $Runspace.ApartmentState = [System.Threading.ApartmentState]::STA
            $Runspace.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::UseCurrentThread
        }
        return $Runspace
    }

    <#
    .SYNOPSIS
    	Title   : Function New-Powershell
    	Author  : Jon Damvi
    	Version : 1.0.0
    	Date    : 01.06.2025
    	License : MIT (LICENSE)
    
    	Release Notes: 
    		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).
    
    .DESCRIPTION
    	Creates new Powershell object.
    
    .PARAMETER Runspace
    	(Mandatory) Specifies Runspace for Powershell object.
    	Expected type: [Runspace]
    
    .INPUTS
    	None
    
    .OUTPUTS
    	System.Management.Automation.PowerShell
    
    .EXAMPLE
    	# Usage Case ... Example Description :
    	PS > New-Powershell -Runspace $Runspace
    
    .LINK
    	GitHub Repository: https://github.com/...
    #>
    Function New-Powershell {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [System.Management.Automation.Runspaces.Runspace]$Runspace
        )
        $Powershell = [powershell]::Create()
        $Powershell.Runspace = $Runspace
        $Powershell.Streams.ClearStreams()
        return $Powershell
    }

    <#
    .SYNOPSIS
    	Title   : Function Add-PowershellScript
    	Author  : Jon Damvi
    	Version : 1.0.0
    	Date    : 01.06.2025
    	License : MIT (LICENSE)
    
    	Release Notes: 
    		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).
    
    .DESCRIPTION
    	Adds scripts to Powershell object.
    
    .PARAMETER Powershell
    	(Mandatory) Specifies Powershell objects to add scripts to.

    	Expected type: [PowerShell]
    
    .PARAMETER Scripts
    	(Mandatory) Specifies scripts to add to Powershell object.
    	Expected type: [Object]
        Supported types: [string], [string[]], [ScriptBlock]
    
    .INPUTS
    	None
    
    .OUTPUTS
    	None
    
    .EXAMPLE
    	# Add one command script to Powershell:
        $Script = 'Write-Host "Test Message"'
        Add-PowershellScript -Powershell $Powershell -Scripts $Script

    .EXAMPLE
    	# Add multiple commands script to Powershell:
        [string[]]$Scripts = @()
        $Scripts += '$TestMessage = "Test Message 1"'
        $Scripts += 'Write-Host $TestMessage'
        Add-PowershellScript -Powershell $Powershell -Scripts $Scripts

    .EXAMPLE
    	# Add Script Block to Powershell and command to execute Script Block:
        $Scripts = {
            $TestMessage = "Test Message 2"
            Write-Host $TestMessage
        }
        Add-PowershellScript -Powershell $Powershell -Scripts $Scripts
    
    .LINK
    	GitHub Repository: https://github.com/...
    #>
    Function Add-PowershellScript {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [powershell]$Powershell,

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [Object]$Scripts
        )
        If ($Scripts -is [ScriptBlock]) {
            [void]$Powershell.AddScript($Scripts.ToString())
        } ElseIf ($Scripts -is [string[]]) {
            Foreach ($Script in $Scripts) {
                [void]$Powershell.AddScript($Script)
            }
        } ElseIf ($Scripts -is [string]) {
            [void]$Powershell.AddScript($Scripts)
        } Else {
            Throw "Invalid input object type '$($Script.GetType().FullName)' for parameter 'Script'. Expected Types: [string], [string[]], [ScriptBlock]."
        }
        
    }

    <#
    .SYNOPSIS
    	Title   : Function Run-Powershell
    	Author  : Jon Damvi
    	Version : 1.0.0
    	Date    : 01.06.2025
    	License : MIT (LICENSE)
    
    	Release Notes: 
    		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).
    
    .DESCRIPTION
    	Executes Powershell asynchronously and waits for completion.
        If Powershell execution exceeds Timeout time, then Powershell
        runspace pipeline will be terminated.
    
    .PARAMETER Powershell
    	(Mandatory) Specifies Powershell object to execute.
    	Expected type: [PowerShell]
    
    .PARAMETER Timeout
    	(Optional) Specifies wait Timeout in seconds.
    	Expected type  : [int]
    	Allowed Values : [1;3600]
    	Default Value  : 60
    
    .INPUTS
    	None
    
    .OUTPUTS
    	$result - result object returned by $Powershell.EndInvoke($asyncResult)
    
    .EXAMPLE
    	Run-Powershell -Powershell $Powershell -Timeout 300
    
    .LINK
    	GitHub Repository: https://github.com/...
    #>
    Function Run-Powershell {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [powershell]$Powershell,

            [Parameter()]
            [ValidateRange(1,3600)]
            [int]$Timeout = 60
        )
        $result = $null
        $TimeoutMS = $Timeout * 1000

        [System.Diagnostics.Stopwatch]$ProcessExecTimer = [System.Diagnostics.Stopwatch]::StartNew()
        $asyncResult = $Powershell.BeginInvoke()

        While(($asyncResult.IsCompleted -ne $true) -And ($ProcessExecTimer.ElapsedMilliseconds -lt $TimeoutMS)) {
            Start-Sleep -Milliseconds 1000
        }
        
        If ($ProcessExecTimer.ElapsedMilliseconds -ge $TimeoutMS) {
            $powershell.Stop()       # Request pipeline stop
            $powershell.Runspace.Close()        # Close runspace (terminates pipeline)
            $powershell.Runspace.Dispose()      # Dispose runspace
            $powershell.Dispose()    # Dispose PowerShell instance
        } Else {
            # Collect all results
            $result = $Powershell.EndInvoke($asyncResult)
        }
        $ProcessExecTimer.Stop() | Out-Null
        return $result
    }

    <#
    .SYNOPSIS
    	Title   : Function Dispose-Powershell
    	Author  : Jon Damvi
    	Version : 1.0.0
    	Date    : 01.06.2025
    	License : MIT (LICENSE)
    
    	Release Notes: 
    		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).
    
    .DESCRIPTION
    	Stops Powershell session object, closes and disposes Runspace and Powershell session object.
    
    .PARAMETER Powershell
    	(Mandatory) Specifies PowerShell session object to dispose.
    	Expected type: [PowerShell]
    
    .INPUTS
    	None
    
    .OUTPUTS
    	None
    
    .EXAMPLE
    	Dispose-Powershell -Powershell $Powershell
    
    .LINK
    	GitHub Repository: https://github.com/...
    #>
    Function Dispose-Powershell {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [powershell]$Powershell
        )

        If ($null -eq $Powershell) { return }

        $powershell.Stop()       # Request pipeline stop
        $powershell.Runspace.Close()        # Close runspace (terminates pipeline)
        $powershell.Runspace.Dispose()      # Dispose runspace
        $powershell.Dispose()    # Dispose PowerShell instance

    }

    <#
    .SYNOPSIS
    Outputs PowerShell runspace execution results with configurable logging levels.

    .DESCRIPTION


    .PARAMETER Result
    The result object returned from PowerShell runspace execution.

    .PARAMETER PowerShell
    The PowerShell instance used for execution.

    .PARAMETER OutputLevel
    Log level (3=Error, 2=Warning, 1=Info, 0=Verbose, -1=Debug, -2=Trace). Default: 1 (Info)

    .PARAMETER DisableOutput
    Suppress standard output stream even if level permits it.

    .PARAMETER DisableInfo  
    Suppress information stream even if level permits it.

    .PARAMETER DisableVerbose
    Suppress verbose stream even if level permits it.

    .PARAMETER DisableDebug
    Suppress debug stream even if level permits it.

    .PARAMETER DisableWarning
    Suppress warning stream even if level permits it.

    .PARAMETER DisableError
    Suppress error stream (rarely needed).

    .EXAMPLE
    Output-RunspaceResult $result $ps
    Shows info, warnings, errors, and output (default level 1).

    .EXAMPLE  
    Output-RunspaceResult $result $ps -OutputLevel 2
    Shows only warnings and errors.

    .EXAMPLE
    Output-RunspaceResult $result $ps -OutputLevel 0 -DisableVerbose
    Shows everything except verbose messages.

    .EXAMPLE
    Output-RunspaceResult $result $ps -OutputLevel -1
    Shows everything including debug messages.
    #>
    <#
    .SYNOPSIS
    	Title   : Function Output-RunspaceResult
    	Author  : Jon Damvi
    	Version : 1.0.0
    	Date    : 01.06.2025
    	License : MIT (LICENSE)
    
    	Release Notes: 
    		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).
    
    .DESCRIPTION
        Displays the various output streams from a PowerShell runspace execution.
        Uses numeric log levels similar to standard logging frameworks:
        - Error   (3): Only errors
        - Warning (2): Errors + warnings  
        - Info    (1): Errors + warnings + information + output [DEFAULT]
        - Verbose (0): All above + verbose messages
        - Debug  (-1): All above + debug messages
        - Trace  (-2): Reserved for future use
    
    .PARAMETER Result
    	(Mandatory) Specifies result object returned from PowerShell runspace execution.
    	Expected type: [Object]
    
    .PARAMETER PowerShell
    	(Mandatory) Specifies PowerShell instance used for execution.
    	Expected type: [PowerShell]
    
    .PARAMETER OutputLevel
    	(Optional) Specifies Log level (3=Error, 2=Warning, 1=Info, 0=Verbose, -1=Debug, -2=Trace). Default: 1 (Info).
    	Expected type: [int]
    	Allowed Values: [-2;3]
    	Default Value: -2 (Trace) - shows everything.
    
    .PARAMETER DisableOutput
    	(Optional) Suppress standard output stream even if level permits it.
    	Expected type: [switch]
    
    .PARAMETER DisableInfo
    	(Optional) Suppress information stream even if level permits it.
    	Expected type: [switch]
    
    .PARAMETER DisableVerbose
    	(Optional) Suppress verbose stream even if level permits it.
    	Expected type: [switch]
    
    .PARAMETER DisableDebug
    	(Optional) Suppress debug stream even if level permits it.
    	Expected type: [switch]
    
    .PARAMETER DisableWarning
    	(Optional) Suppress warning stream even if level permits it.
    	Expected type: [switch]
    
    .PARAMETER DisableError
    	(Optional) Suppress error stream (rarely needed).
    	Expected type: [switch]
    
    .INPUTS
    	None
    
    .OUTPUTS
    	None
    
    .EXAMPLE
        # Shows info, warnings, errors, and output (default level 1).
        Output-RunspaceResult $result $Powershell

    .EXAMPLE
        # Shows only warnings and errors.
        Output-RunspaceResult $result $Powershell -OutputLevel 2

    .EXAMPLE
        # Shows everything except verbose messages.
        Output-RunspaceResult $result $Powershell -OutputLevel 0 -DisableVerbose

    .EXAMPLE
        # Shows everything including debug messages.
        Output-RunspaceResult $result $Powershell -OutputLevel -1

    .LINK
    	GitHub Repository: https://github.com/...
    #>
    Function Output-RunspaceResult {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory, Position = 0)]
            [Object]$Result,

            [Parameter(Mandatory, Position = 1)]
            [PowerShell]$PowerShell,

            # Log levels: Error=3, Warning=2, Info=1, Verbose=0, Debug=-1, Trace=-2
            [ValidateRange(-2, 3)]
            [int]$OutputLevel = -2,  # Default: Info level
        
            # Fine-grain control overrides for specific streams
            [switch]$DisableOutput,      # Disable standard output stream
            [switch]$DisableInfo,        # Disable information stream  
            [switch]$DisableVerbose,     # Disable verbose stream
            [switch]$DisableDebug,       # Disable debug stream
            [switch]$DisableWarning,     # Disable warning stream
            [switch]$DisableError        # Disable error stream (rarely used)
        )
    
        If ($null -eq $Result) { return }

        # Determine what streams to show based on output level and overrides
        $showOutput = ($OutputLevel -le 1) -and -not $DisableOutput      # Info and below
        $showInfo = ($OutputLevel -le 1) -and -not $DisableInfo          # Info and below  
        $showVerbose = ($OutputLevel -le 0) -and -not $DisableVerbose    # Verbose and below
        $showDebug = ($OutputLevel -le -1) -and -not $DisableDebug       # Debug and below
        $showWarning = ($OutputLevel -le 2) -and -not $DisableWarning    # Warning and below
        $showError = ($OutputLevel -le 3) -and -not $DisableError        # Always (unless explicitly disabled)

        # Show header unless completely silent
        if ($OutputLevel -le 2) {
            $levelName = switch ($OutputLevel) {
                 3 { "ERROR"   }
                 2 { "WARNING" }  
                 1 { "INFO"    }
                 0 { "VERBOSE" }
                -1 { "DEBUG"   }
                -2 { "TRACE"   }
            }
            Write-Host "=== RUNSPACE EXECUTION RESULTS (Level: $levelName) ===" -ForegroundColor Green
        }

        # Standard Output Stream (Write-Output, return values)
        if ($showOutput -and $Result -and $Result.Count -gt 0) {
            Write-Host "Output Stream:" -ForegroundColor Cyan
            $filteredOutput = $Result | Where-Object { 
                $_ -ne $null -and ![string]::IsNullOrWhiteSpace($_.ToString()) 
            }
            if ($filteredOutput) {
                $filteredOutput | % {
                    Write-Host "  $($_)" -ForegroundColor White
                }
            } else {
                Write-Host "  (No meaningful output)" -ForegroundColor Gray
            }
        }

        # PowerShell Output Stream (alternative access method)
        if ($showOutput) {
            $psOutput = $PowerShell.Streams.Output
            if ($psOutput -and $psOutput.Count -gt 0) {
                Write-Host "PowerShell Output Stream:" -ForegroundColor Cyan
                $psOutput | % { 
                    Write-Host "  $($_)" -ForegroundColor White
                }
            }
        }

        # Information Stream (Write-Information, Write-Host in PS5+)
        if ($showInfo -and $PowerShell.Streams.Information.Count -gt 0) {
            Write-Host "Information Stream:" -ForegroundColor Cyan
            $PowerShell.Streams.Information | % { 
                Write-Host "  $($_.MessageData)" -ForegroundColor White 
            }
        }

        # Verbose Stream (Write-Verbose)
        if ($showVerbose -and $PowerShell.Streams.Verbose.Count -gt 0) {
            Write-Host "Verbose Stream:" -ForegroundColor Cyan
            $PowerShell.Streams.Verbose | % { 
                Write-Host "  $($_.Message)" -ForegroundColor Gray
            }
        }

        # Debug Stream (Write-Debug)
        if ($showDebug -and $PowerShell.Streams.Debug.Count -gt 0) {
            Write-Host "Debug Stream:" -ForegroundColor Cyan
            $PowerShell.Streams.Debug | % { 
                Write-Host "  $($_.Message)" -ForegroundColor DarkGray
            }
        }

        # Warning Stream (Write-Warning)
        if ($showWarning -and $PowerShell.Streams.Warning.Count -gt 0) {
            Write-Host "Warning Stream:" -ForegroundColor Yellow
            $PowerShell.Streams.Warning | % { 
                Write-Host "  WARNING: $($_.Message)" -ForegroundColor Yellow
            }
        }

        # Error Stream (Write-Error, exceptions)
        if ($showError -and $PowerShell.Streams.Error.Count -gt 0) {
            Write-Host "Error Stream:" -ForegroundColor Red
            $PowerShell.Streams.Error | % { 
                Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
                if ($OutputLevel -le 0) {  # Show stack trace at Verbose level and below
                    Write-Host "    $($_.ScriptStackTrace)" -ForegroundColor DarkRed
                }
            }
        }
    }

    <#
    .SYNOPSIS
    	Title   : Function Open-Runspace
    	Author  : Jon Damvi
    	Version : 1.0.0
    	Date    : 01.06.2025
    	License : MIT (LICENSE)
    
    	Release Notes: 
    		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).
    
    .DESCRIPTION
    	Opens Runspace.
    
    .PARAMETER Runspace
    	(Mandatory) Specifies Runspace to open.
    	Expected type: [Runspace]
    
    .INPUTS
    	None
    
    .OUTPUTS
    	None
    
    .EXAMPLE
    	# Open Runspace:
    	Open-Runspace -Runspace $Runspace
    
    .LINK
    	GitHub Repository: https://github.com/...
    #>
    Function Open-Runspace {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [System.Management.Automation.Runspaces.Runspace]$Runspace
        )
        $Runspace.Open()
    }

    <#
    .SYNOPSIS
    	Title   : Function Add-RunspaceVariable
    	Author  : Jon Damvi
    	Version : 1.0.0
    	Date    : 01.06.2025
    	License : MIT (LICENSE)
    
    	Release Notes: 
    		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).
    
    .DESCRIPTION
    	Adds specified variable to Runspace.
    
    .PARAMETER Runspace
    	(Mandatory) Specifies Runspace to add variable to.
    	Expected type: [Runspace]
    
    .PARAMETER VariableName
    	(Mandatory) Specifies variable name.
    	Expected type: [String]
    
    .PARAMETER Variable
    	(Mandatory) Specifies variable value.
    	Expected type: [Object]
    
    .INPUTS
    	None
    
    .OUTPUTS
    	None
    
    .EXAMPLE
    	# Add variable to Runspace:
        $Variable = "Test String"
    	Add-RunspaceVariable -Runspace $Runspace -VariableName "Variable" -Variable $Variable
    
    .LINK
    	GitHub Repository: https://github.com/...
    #>
    Function Add-RunspaceVariable {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [System.Management.Automation.Runspaces.Runspace]$Runspace,

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$VariableName,

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [Object]$Variable
        )
        $Runspace.SessionStateProxy.SetVariable($VariableName, $Variable)
    }

    <#
    .SYNOPSIS
    	Title   : Function Execute-InRunspace
    	Author  : Jon Damvi
    	Version : 1.0.0
    	Date    : 01.06.2025
    	License : MIT (LICENSE)
    
    	Release Notes: 
    		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).
    
    .DESCRIPTION
    	Includes specified variables in Runspace and executes specified script in Runspace.
    
    .PARAMETER Script
    	(Mandatory) Specifies script to execute in New Runspace.
    	Expected type   : [Object]
        Supported types : [string], [string[]], [ScriptBlock]
    
    .PARAMETER Variables
    	(Optional) Specifies variables to include in New Runspace.
    	Expected type: [PSVariable[]]
    
    .INPUTS
    	None
    
    .OUTPUTS
    	None
    
    .EXAMPLE
    	# Execute Write-Host in New Runspace and output variable value:
        $Variable = "Test Message"
        $Script = { Write-Host $Variable }
        Execute-InRunspace -Script $Script -Variables @(Get-Variable "Variable")
    
    .LINK
    	GitHub Repository: https://github.com/...
    #>
    Function Execute-InRunspace {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [Object]$Script,

            [System.Management.Automation.PSVariable[]]$Variables
        )

        $runspace = New-Runspace -Interactive

        Open-Runspace $runspace

        If($Variables) {
            Foreach ($Variable in $Variables) {
                Add-RunspaceVariable $runspace $Variable.Name $Variable.Value
            }
        }
        
        $powershell = New-Powershell $runspace

        Add-PowershellScript $powershell $Script

        $result = Run-Powershell $powershell

        Output-RunspaceResult $result $powershell

        Dispose-Powershell $powershell
    }
}

$NotificationParams = @{
    SourceName = "MyApp"
    Title      = "Network Drives:"
    Message    = "Drives have been mapped successfully"
    #Muted      = $false
    IconPath   = "C:\Windows\System32\wpcatltoast.png"
    AudioPath  = "C:\Windows\Media\Windows Pop-up Blocked.wav"
    Type = 'default'
    ExpirationTime = [DateTimeOffset]::Now.AddMinutes(3)
}


[string[]]$Scripts = @()

$Scripts += $NotificationFunctionsDefs.ToString()
$Scripts += "Show-Notification @NotificationParams -Verbose"

$Variables = @(Get-Variable 'NotificationParams')

. $RunspaceFunctionsDefs

Execute-InRunspace $Scripts $Variables
