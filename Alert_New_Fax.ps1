#*************************************Alert_New_Fax.ps1*******************************#
#*********************************|Script by: Stephen Nelson|*************************#
#*********************************|Script created: 12/01/19|**************************#
#Purpose:
#This script was pieced together using the above source and customized to suite a 
#specific need I have for alerting a user that a new fax has been recieved. This script
#watches a directory and creates an event ID which Task Scheduler is set to watch for. 
#When the event ID is detected by Task Scheduler, it notifies the user in the action
#center and gives a balloon notification.
#=====================================================================================#

#<Set Path to be monitored>#

#INTAKE
$searchPath = "\\CCSOFS1\CCDF_Fax\226-5236 Intake Fax #1"
$searchPath1 = "\\CCSOFS1\CCDF_Fax\226-5256 Intake Fax #2"
$searchPath2 = "\\CCSOFS1\CCDF_Fax\226-5254 New Booking Fax"
$searchPath3 = "\\CCSOFS1\CCDF_Fax\226-5345 Jail Intake Lt. Fax"
#MEDICAL
$searchPath4 = "\\CCSOFS1\CCDF_Fax\226-5228 Medical Fax"
#COURT OFFICE
$searchPath5 = "\\CCSOFS1\CCDF_Fax\226-5297 Court Sgt. Fax"
#HOUSING
$searchPath6 = "\\CCSOFS1\CCDF_Fax\226-5292 A Tower Fax"
$searchPath7 = "\\CCSOFS1\CCDF_Fax\226-5295 B Tower Fax"
$searchPath8 = "\\CCSOFS1\CCDF_Fax\226-5333 C Tower Fax"
$searchPath9 = "\\CCSOFS1\CCDF_Fax\226-5283 F Dorm Fax"
$searchPath10 = "\\CCSOFS1\CCDF_Fax\226-5296 Jail Housing Lt. Fax"
#SUPPORT SQUAD
$searchPath11 = "\\CCSOFS1\CCDF_Fax\226-5279 Det. Liason Officer Fax"
$searchPath12 = "\\CCSOFS1\CCDF_Fax\226-5298 Commissary Fax"
$searchPath13 = "\\CCSOFS1\CCDF_Fax\226-5212 Admin LT Fax"
#OTHER
$searchPath14 = "\\CCSOFS1\CCDF_Fax\226-5277 Maintenance Fax"
$searchPath15 = "\\CCSOFS1\CCDF_Fax\226-5299 Jail Library Fax"

#<Set Watch routine>#

#INTAKE
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $searchPath
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true

$watcher1 = New-Object System.IO.FileSystemWatcher
$watcher1.Path = $searchPath1
$watcher1.IncludeSubdirectories = $false
$watcher1.EnableRaisingEvents = $true

$watcher2 = New-Object System.IO.FileSystemWatcher
$watcher2.Path = $searchPath2
$watcher2.IncludeSubdirectories = $false
$watcher2.EnableRaisingEvents = $true

$watcher3 = New-Object System.IO.FileSystemWatcher
$watcher3.Path = $searchPath3
$watcher3.IncludeSubdirectories = $false
$watcher3.EnableRaisingEvents = $true

#MEDICAL

$watcher4 = New-Object System.IO.FileSystemWatcher
$watcher4.Path = $searchPath4
$watcher4.IncludeSubdirectories = $false
$watcher4.EnableRaisingEvents = $true

$watcher5 = New-Object System.IO.FileSystemWatcher
$watcher5.Path = $searchPath5
$watcher5.IncludeSubdirectories = $false
$watcher5.EnableRaisingEvents = $true

$watcher6 = New-Object System.IO.FileSystemWatcher
$watcher6.Path = $searchPath6
$watcher6.IncludeSubdirectories = $false
$watcher6.EnableRaisingEvents = $true

$watcher7 = New-Object System.IO.FileSystemWatcher
$watcher7.Path = $searchPath7
$watcher7.IncludeSubdirectories = $false
$watcher7.EnableRaisingEvents = $true

$watcher8 = New-Object System.IO.FileSystemWatcher
$watcher8.Path = $searchPath8
$watcher8.IncludeSubdirectories = $false
$watcher8.EnableRaisingEvents = $true

$watcher9 = New-Object System.IO.FileSystemWatcher
$watcher9.Path = $searchPath9
$watcher9.IncludeSubdirectories = $false
$watcher9.EnableRaisingEvents = $true

$watcher10 = New-Object System.IO.FileSystemWatcher
$watcher10.Path = $searchPath10
$watcher10.IncludeSubdirectories = $false
$watcher10.EnableRaisingEvents = $true

$watcher11 = New-Object System.IO.FileSystemWatcher
$watcher11.Path = $searchPath11
$watcher11.IncludeSubdirectories = $false
$watcher11.EnableRaisingEvents = $true

$watcher12 = New-Object System.IO.FileSystemWatcher
$watcher12.Path = $searchPath12
$watcher12.IncludeSubdirectories = $false
$watcher12.EnableRaisingEvents = $true

$watcher13 = New-Object System.IO.FileSystemWatcher
$watcher13.Path = $searchPath13
$watcher13.IncludeSubdirectories = $false
$watcher13.EnableRaisingEvents = $true

$watcher14 = New-Object System.IO.FileSystemWatcher
$watcher14.Path = $searchPath14
$watcher14.IncludeSubdirectories = $false
$watcher14.EnableRaisingEvents = $true

$watcher15 = New-Object System.IO.FileSystemWatcher
$watcher15.Path = $searchPath15
$watcher15.IncludeSubdirectories = $false
$watcher15.EnableRaisingEvents = $true

#<Set Events>#

#$changed = Register-ObjectEvent $watcher "Changed" -Action {
#   write-host "Changed: $($eventArgs.FullPath)"
#}

$created = Register-ObjectEvent $watcher "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in Intake Fax #1'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}
$created = Register-ObjectEvent $watcher1 "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in Intake Fax #2'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}
$created = Register-ObjectEvent $watcher2 "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in New Booking Fax'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}
$created = Register-ObjectEvent $watcher3 "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in Jail Intake Lt. Fax'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}
$created = Register-ObjectEvent $watcher4 "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in Medical Fax'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}
$created = Register-ObjectEvent $watcher5 "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in Court Sgt. Fax'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}
$created = Register-ObjectEvent $watcher6 "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in A Tower Fax'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}
$created = Register-ObjectEvent $watcher7 "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in B Tower Fax'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}
$created = Register-ObjectEvent $watcher8 "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in C Tower Fax'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}
$created = Register-ObjectEvent $watcher9 "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in F Dorm Fax'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}
$created = Register-ObjectEvent $watcher10 "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in Jail Housing Lt. Fax'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}
$created = Register-ObjectEvent $watcher11 "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in Det. Liason Officer Fax'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}
$created = Register-ObjectEvent $watcher12 "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in Commissary Fax'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}
$created = Register-ObjectEvent $watcher13 "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in Admin LT Fax'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}
$created = Register-ObjectEvent $watcher14 "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in Maintenance Fax'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}
$created = Register-ObjectEvent $watcher15 "Created" -Action {
    Add-Type -AssemblyName System.Windows.Forms
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::info
    $balloon.BalloonTipText = 'A new fax has been received in Jail Library Fax'
    $balloon.BalloonTipTitle = "Attention $Env:USERNAME"
    $balloon.Visible = $true
    $balloon.ShowBalloonTip(10000)
}

#$deleted = Register-ObjectEvent $watcher "Deleted" -Action {
#   write-host "Deleted: $($eventArgs.FullPath)"
#}
#$renamed = Register-ObjectEvent $watcher "Renamed" -Action {
#   write-host "Renamed: $($eventArgs.FullPath)"
#}