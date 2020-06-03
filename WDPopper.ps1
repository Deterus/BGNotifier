using namespace Windows.Storage
using namespace Windows.Graphics.Imaging

##########################################
### CHANGE THESE SETTINGS TO YOUR OWN! ###
##########################################


### REQUIRED SETTINGS ###
#########################

# This is currently built to only work with discord.

## DISCORD ##
$discord = $True
# Discord Channel Webhook.
$discordWebHook = "https://discordapp.com/api/webhooks/717568200665071636/QNWANbApXesshGhuDgAA5fO6oaweK26qna51dStRk3pBlwtMK1bg5giWzNtRuLvrLtLn"

### OPTIONAL ADVANCED SETTINGS ###
##################################

# Coordinates of Unit Scan window.
$useMyOwnCoordinates = "Yes" # Don't Change This.
$topleftX     = 533
$topLeftY     = 649
$bottomRightX = 1328
$bottomRightY = 922

# Screenshot Location to save temporary img to for OCR Scan. Change if you want it somewhere else.
$path = $env:temp

# Amount of seconds to wait before scanning for a Unit Scan window.
# Note: this script uses hardly any resources and is very quick at the screenshot/OCR process.
$delay = 5

# Option to stop WDPopper once a WD has popped. "Yes" to stop the program, or "No" to keep it running.
# Default is 'Yes', stop scanning after it detects a WD pop.
$stopOnWDPop = "Yes"

# Option on the Number of times to Send the Message to Discord
$timesToPing = 5

# Option to set the delay between pings on discord in seconds
$pingDelay = 2

# Option to notify if you get disconnected. Looks for the 'disconnected' message on the login screen.
$notifyOnDisconnect = $False

#########################################
### DO NOT MODIFY ANYTHING BELOW THIS ###
#########################################

# Force tls1.2 - mainly for telegram since they recently changed this in FEB2020
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Screenshot method
Add-Type -AssemblyName System.Windows.Forms,System.Drawing

# Add the WinRT assembly, and load the appropriate WinRT types
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile,                Windows.Storage,         ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine,                Windows.Foundation,      ContentType = WindowsRuntime]
$null = [Windows.Foundation.IAsyncOperation`1,       Windows.Foundation,      ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap,    Windows.Foundation,      ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.RandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]

# used to find the World Dragon popup location coordinates on your monitor
function Get-Coords {

    $form.TopMost = $True
    $script:label_coords_text.Enabled = $False
    $script:label_coords_text.Visible = $False
    $button_start.Visible = $False
    $textBox.Visible = $False
    $label_status.Text = ""
    $label_status.Refresh()
    $script:label_coords1.Text = ""
    $script:label_coords1.Refresh()
    $script:label_coords2.Text = ""
    $script:label_coords2.Refresh()
    $script:label_coords1.Visible = $True
    $script:label_coords2.Visible = $True
    $script:label_coords_text2.Visible = $True
    $script:label_coords_text2.Enabled = $True
    $script:cancelLoop = $False
    $count = 1

    :coord While( $true ) {
        
        If( (([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftShift)) -and ([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftCtrl))) -or ($script:cancelLoop) -or ($count -ge 3)) { 
            Break
        }
        If( [System.Windows.Forms.UserControl]::MouseButtons -ne "None" ) { 
          While( [System.Windows.Forms.UserControl]::MouseButtons -ne "None" ) {
            Start-Sleep -Milliseconds 100 # Wait for the MOUSE UP event
            [System.Windows.Forms.Application]::DoEvents()
          }
        
            $mp = [Windows.Forms.Cursor]::Position

            if ($count -eq 1) {
                $script:label_coords1.Text = "Top left: $($mp.ToString().Replace('{','').Replace('}',''))" 
                $script:label_coords1.Refresh()
                $count++
            }
            elseif ($count -eq 2) {
                $script:label_coords2.Text = "Bottom Right: $($mp.ToString().Replace('{','').Replace('}',''))"
                $script:label_coords2.Refresh()
                $count++
            }
            if ($count -ge 3) {
                Break coord
            }           
            
        }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100        

    }
    #[System.Windows.Forms.Application]::DoEvents()
    if (($script:cancelLoop) -or ($count -ge 3)) {
        Return
    }    
}

# Screenshot function
function Get-WDPop {

    $bounds   = [Drawing.Rectangle]::FromLTRB($topleftX, $topLeftY, $bottomRightX, $bottomRightY)
    $pic      = New-Object System.Drawing.Bitmap ([int]$bounds.width), ([int]$bounds.height)
    $graphics = [Drawing.Graphics]::FromImage($pic)

    $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size)

    $pic.Save("$path\WDPopper_Img.png")

    $graphics.Dispose()
    $pic.Dispose()

}

# OCR Scan Function
function Get-Ocr {

# Takes a path to an image file, with some text on it.
# Runs Windows 10 OCR against the image.
# Returns an [OcrResult], hopefully with a .Text property containing the text
# OCR part of the script from: https://github.com/HumanEquivalentUnit/PowerShell-Misc/blob/master/Get-Win10OcrTextFromImage.ps1


    [CmdletBinding()]
    Param
    (
        # Path to an image file
        [Parameter(Mandatory=$true, 
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true, 
                    Position=0,
                    HelpMessage='Path to an image file, to run OCR on')]
        [ValidateNotNullOrEmpty()]
        $Path
    )

    Begin {
    
        # [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages
        $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    

        # PowerShell doesn't have built-in support for Async operations, 
        # but all the WinRT methods are Async.
        # This function wraps a way to call those methods, and wait for their results.
        $getAwaiterBaseMethod = [WindowsRuntimeSystemExtensions].GetMember('GetAwaiter').
                                    Where({
                                            $PSItem.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
                                        }, 'First')[0]

        Function Await {
            param($AsyncTask, $ResultType)

            $getAwaiterBaseMethod.
                MakeGenericMethod($ResultType).
                Invoke($null, @($AsyncTask)).
                GetResult()
        }
    }

    Process
    {
        foreach ($p in $Path)
        {
      
            # From MSDN, the necessary steps to load an image are:
            # Call the OpenAsync method of the StorageFile object to get a random access stream containing the image data.
            # Call the static method BitmapDecoder.CreateAsync to get an instance of the BitmapDecoder class for the specified stream. 
            # Call GetSoftwareBitmapAsync to get a SoftwareBitmap object containing the image.
            #
            # https://docs.microsoft.com/en-us/windows/uwp/audio-video-camera/imaging#save-a-softwarebitmap-to-a-file-with-bitmapencoder

            # .Net method needs a full path, or at least might not have the same relative path root as PowerShell
            $p = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p)
        
            $params = @{ 
                AsyncTask  = [StorageFile]::GetFileFromPathAsync($p)
                ResultType = [StorageFile]
            }
            $storageFile = Await @params


            $params = @{ 
                AsyncTask  = $storageFile.OpenAsync([FileAccessMode]::Read)
                ResultType = [Streams.IRandomAccessStream]
            }
            $fileStream = Await @params


            $params = @{
                AsyncTask  = [BitmapDecoder]::CreateAsync($fileStream)
                ResultType = [BitmapDecoder]
            }
            $bitmapDecoder = Await @params


            $params = @{ 
                AsyncTask = $bitmapDecoder.GetSoftwareBitmapAsync()
                ResultType = [SoftwareBitmap]
            }
            $softwareBitmap = Await @params

            # Run the OCR
            Await $ocrEngine.RecognizeAsync($softwareBitmap) ([Windows.Media.Ocr.OcrResult])

        }
    }
}

# get window and sizes function
Function Get-Window {
    <#
        .NOTES
            Name: Get-Window
            Author: Boe Prox
    #>
    [OutputType('System.Automation.WindowInfo')]
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipelineByPropertyName=$True)]
        $ProcessName
    )
    Begin {
        Try{
            [void][Window]
        } Catch {
        Add-Type @"
              using System;
              using System.Runtime.InteropServices;
              public class Window {
                [DllImport("user32.dll")]
                [return: MarshalAs(UnmanagedType.Bool)]
                public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
              }
              public struct RECT
              {
                public int Left;        // x position of upper-left corner
                public int Top;         // y position of upper-left corner
                public int Right;       // x position of lower-right corner
                public int Bottom;      // y position of lower-right corner
              }
"@
        }
    }
    Process {        
        Get-Process -Name $ProcessName | ForEach {
            $Handle = $_.MainWindowHandle
            $Rectangle = New-Object RECT
            $Return = [Window]::GetWindowRect($Handle,[ref]$Rectangle)
            If ($Return) {
                $Height = $Rectangle.Bottom - $Rectangle.Top
                $Width = $Rectangle.Right - $Rectangle.Left
                $Size = New-Object System.Management.Automation.Host.Size -ArgumentList $Width, $Height
                $TopLeft = New-Object System.Management.Automation.Host.Coordinates -ArgumentList $Rectangle.Left, $Rectangle.Top
                $BottomRight = New-Object System.Management.Automation.Host.Coordinates -ArgumentList $Rectangle.Right, $Rectangle.Bottom
                If ($Rectangle.Top -lt 0 -AND $Rectangle.LEft -lt 0) {
                    Write-Warning "Window is minimized! Coordinates will not be accurate."
                }
                $Object = [pscustomobject]@{
                    ProcessName = $ProcessName
                    Size = $Size
                    TopLeft = $TopLeft
                    BottomRight = $BottomRight
                }
                $Object.PSTypeNames.insert(0,'System.Automation.WindowInfo')
                $Object
            }
        }
    }
}

# default screenshot area if no coordinates specified in the above user section.
# Also tries to detect which window your game is running on, if using multiple monitors
# Get's the middlle top half of the screen area to look for BG Queue pop and disconnect messages
if ($useMyOwnCoordinates -eq "No") {
    $window = Get-Process | ? {$_.MainWindowTitle -like "World of Warcraft"} | Get-Window | select -First 1
    $topleftX = [math]::floor($window.BottomRight.x / 3)
    $topLeftY = 0
    $bottomRightX = [math]::floor($topLeftX * 2)
    $bottomRightY = [math]::floor($window.BottomRight.y / 2)
}


# Notification function
function WDPopper {
    $disconnected = $False
    $button_start.Enabled = $False
    $button_start.Visible = $False
    $textBoxCurrValue = $textBox.Text
    $textBox.Visible = $False
    $button_stop.Enabled = $True
    $button_stop.Visible = $True
    $form.MinimizeBox = $False # disable while running since it breaks things
    $script:label_coords_text.Visible = $False
    $label_status.ForeColor = "#7CFC00"
    $label_status.text = "WDPopper is Running!"
    $label_status.Refresh()
    $script:cancelLoop = $False

    :check Do {
        # check for clicks in the form since we are looping
        for ($i=0; $i -lt $delay; $i++) {

            [System.Windows.Forms.Application]::DoEvents()

            if ($script:cancelLoop) {
                $button_start.Enabled = $True
                $button_start.Visible = $True
                $textBox.Visible = $True
                $button_stop.Enabled = $False
                $button_stop.Visible = $False
                $form.MinimizeBox = $True
                $label_status.text = ""
                $label_status.Refresh()
                $script:label_coords_text.Visible = $True
                Break check
            }

            Sleep -Seconds 1
        }

        Get-WDPop
        $wdAlert = (Get-Ocr $path\WDPopper_Img.png).Text
        if ($notifyOnDisconnect) {
            if ($wdAlert -like "*disconnected*") {
                $disconnected = $True
            }
        }     
    }
    Until (($wdAlert -like "*Azuregos*") -or ($wdAlert -like "*Emeriss*") -or ($wdAlert -like "*Lethon*") -or ($wdAlert -like "*Lord Kazzak*") -or ($wdAlert -like "*Taerar*") -or ($wdAlert -like "*Ysondre*") -or ($wdAlert -like "*Elder Mottled*") -or ($disconnected))

    if ($script:cancelLoop) {
        Return
    }

    # set messages
    if ($wdAlert -like "*Azuregos*") {
        $msg = "The Mighty Azuregos Has Spawned! @here @everyone Specific Info:" + $textBoxCurrValue
    }
    elseif ($wdAlert -like "*Emeriss*") {
        $msg = "The Powerful Emeriss Has Spawned! @here @everyone Specific Info:" + $textBoxCurrValue
    }
    elseif ($wdAlert -like "*Lethon*") {
        $msg = "The Dreadful Lethon Has Spawned! @here @everyone Specific Info:" + $textBoxCurrValue
    }
    elseif ($wdAlert -like "*Lord Kazzak*") {
        $msg = "The Dark Lord Kazzak Has Spawned! @here @everyone Specific Info:" + $textBoxCurrValue
    }
    elseif ($wdAlert -like "*Taerar*") {
        $msg = "The Terrible Taerar Has Spawned! @here @everyone Specific Info:" + $textBoxCurrValue
    }
    elseif ($wdAlert -like "*Ysondre*") {
        $msg = "The Baleful Ysondre Has Spawned! @here @everyone Specific Info:" + $textBoxCurrValue
    }
    elseif ($wdAlert -like "*disconnected*") {
        $msg = "You've been Disconnected!"
    }
    # elseif ($wdAlert -like "*Elder Mottled*") {
    #    $msg = "Your Elder Boar has Popped Joe is our new king. Specific Info:" + $textBoxCurrValue
    # }
    

    # msg Discord
    if ($discord) {

        $discordHeaders = @{
            "Content-Type" = "application/json"
        }

        $discordBody = @{
            content = $msg
        } | convertto-json

        if($stopOnWDPop -and $timesToPing){
            For ($i=0; $i -lt $timesToPing; $i++) {
                Invoke-RestMethod -Uri $discordWebHook -Method POST -Headers $discordHeaders -Body $discordBody
                Start-Sleep -Seconds $pingDelay
            }
        } else {
            Invoke-RestMethod -Uri $discordWebHook -Method POST -Headers $discordHeaders -Body $discordBody
        }

    }
    
    if ($wdAlert -like "*disconnected*") {
        $label_status.ForeColor = "#FFFF00"
        $label_status.text = "You've been Disconnected!"
        $label_status.Refresh()
        $button_stop.Enabled = $False
        $button_stop.Visible = $False
        $button_start.Enabled = $True
        $button_start.Visible = $True
        $textBox.Visible = $True
        $script:label_coords_text.Visible = $True
        $form.MinimizeBox = $True
    }
    elseif ($stopOnWDPop -eq "Yes") {
        $label_status.ForeColor = "#FFFF00"
        $label_status.text = "Your World Dragon Spawned!"
        $label_status.Refresh()
        $button_stop.Enabled = $False
        $button_stop.Visible = $False
        $button_start.Enabled = $True
        $button_start.Visible = $True
        $textBox.Visible = $True
        $script:label_coords_text.Visible = $True
        $form.MinimizeBox = $True
    }
    elseif ($stopOnWDPop -eq "No") {
        WDPopper
    }
}

# Form section
$form                           = New-Object System.Windows.Forms.Form
$form.Text                      ='WDPopper'
$form.Width                     = 250
$form.Height                    = 175
$form.AutoSize                  = $True
$form.MaximizeBox               = $False
$form.BackColor                 = "#4a4a4a"
$form.TopMost                   = $False
$form.StartPosition             = 'CenterScreen'
$form.FormBorderStyle           = "FixedDialog"

# Start Button
$button_start                   = New-Object system.Windows.Forms.Button
$button_start.BackColor         = "#f5a623"
$button_start.text              = "START"
$button_start.width             = 200
$button_start.height            = 50
$button_start.location          = New-Object System.Drawing.Point(30,7)
$button_start.Font              = 'Microsoft Sans Serif,9,style=Bold'
$button_start.FlatStyle         = "Flat"

# Stop Button
$button_stop                    = New-Object system.Windows.Forms.Button
$button_stop.BackColor          = "#f5a623"
$button_stop.ForeColor          = "#FF0000"
$button_stop.text               = "STOP"
$button_stop.width              = 120
$button_stop.height             = 50
$button_stop.location           = New-Object System.Drawing.Point(62,15)
$button_stop.Font               = 'Microsoft Sans Serif,9,style=Bold'
$button_stop.FlatStyle          = "Flat"
$button_stop.Enabled            = $False
$button_stop.Visible            = $False

$textBox                        = New-Object System.Windows.Forms.TextBox
$textBox.Location               = New-Object System.Drawing.Point(10,75)
$textBox.Size                   = New-Object System.Drawing.Size(260,20)
$textBoxCurrValue               = ""

# Status label
$label_status                   = New-Object system.Windows.Forms.Label
$label_status.text              = ""
$label_status.AutoSize          = $True
$label_status.width             = 30
$label_status.height            = 20
$label_status.location          = New-Object System.Drawing.Point(60,75)
$label_status.Font              = 'Microsoft Sans Serif,10,style=Bold'
$label_status.ForeColor         = "#7CFC00"

# Coords label text
$script:label_coords_text            = New-Object system.Windows.Forms.LinkLabel
$script:label_coords_text.text       = "Get Coords"
$script:label_coords_text.AutoSize   = $True
$script:label_coords_text.width      = 30
$script:label_coords_text.height     = 20
$script:label_coords_text.location   = New-Object System.Drawing.Point(5,100)
$script:label_coords_text.Font       = 'Microsoft Sans Serif,9,'
$script:label_coords_text.ForeColor  = "#00ff00"
$script:label_coords_text.LinkColor  = "#f5a623"
$script:label_coords_text.ActiveLinkColor = "#f5a623"
$script:label_coords_text.add_Click({Get-Coords})

# Coords label text exit
$script:label_coords_text2            = New-Object system.Windows.Forms.LinkLabel
$script:label_coords_text2.text       = "Exit Coords"
$script:label_coords_text2.AutoSize   = $True
$script:label_coords_text2.width      = 30
$script:label_coords_text2.height     = 20
$script:label_coords_text2.location   = New-Object System.Drawing.Point(5,100)
$script:label_coords_text2.Font       = 'Microsoft Sans Serif,9,'
$script:label_coords_text2.ForeColor  = "#00ff00"
$script:label_coords_text2.LinkColor  = "#f5a623"
$script:label_coords_text2.ActiveLinkColor = "#f5a623"
$script:label_coords_text2.Visible    = $False
$script:label_coords_text2.add_Click({
    $script:cancelLoop = $True
    $script:label_coords1.Visible = $False
    $script:label_coords2.Visible = $False
    $button_start.Visible = $True
    $textBox.Visible = $True
    $script:label_coords1.Text = ""
    $script:label_coords1.Refresh()
    $script:label_coords2.Text = ""
    $script:label_coords2.Refresh()
    $script:label_coords_text2.Visible = $False
    $script:label_coords_text.Visible = $True
    $script:label_coords_text.Enabled = $True
    $script:label_coords_text2.Enabled = $False
    $form.TopMost = $False
})

# Coords label top left
$script:label_coords1            = New-Object system.Windows.Forms.Label
$script:label_coords1.Text       = ""
$script:label_coords1.AutoSize   = $True
$script:label_coords1.width      = 30
$script:label_coords1.height     = 20
$script:label_coords1.location   = New-Object System.Drawing.Point(10,15)
$script:label_coords1.Font       = 'Microsoft Sans Serif,10,style=Bold'
$script:label_coords1.ForeColor  = "#f5a623"

# Coords label bottom right
$script:label_coords2            = New-Object system.Windows.Forms.Label
$script:label_coords2.Text       = ""
$script:label_coords2.AutoSize   = $True
$script:label_coords2.width      = 30
$script:label_coords2.height     = 20
$script:label_coords2.location   = New-Object System.Drawing.Point(10,40)
$script:label_coords2.Font       = 'Microsoft Sans Serif,10,style=Bold'
$script:label_coords2.ForeColor  = "#f5a623"

# add all controls
$form.Controls.AddRange(($button_start,$button_stop,$textBox,$label_status,$script:label_coords_text,$script:label_coords_text2,$script:label_coords1,$script:label_coords2))

# Button methods
$button_start.Add_Click({WDPopper})
$button_stop.Add_Click({
    if (Test-Path $path\WDPopper_Img.png) {
        Remove-Item $path\WDPopper_Img.png -Force -Confirm:$False
    }
    $script:cancelLoop = $True
})

# catch close handle
$form.add_FormClosing({
    if (Test-Path $path\WDPopper_Img.png) {
        Remove-Item $path\WDPopper_Img.png -Force -Confirm:$False
    }
    $script:cancelLoop = $True
})

# show the forms
$form.ShowDialog()

# close the forms
$form.Dispose()