
# Constants (these will need to be maintained)
$exePath = "C:\Program Files\crunchlabs-arduino-create-agent\crunchlabs-arduino-create-agent.exe"
$crunchlabsAgentURL = "https://ide.crunchlabs.com/assets/downloads/windows/crunchlabs-arduino-create-agent%20Setup%201.0.2.exe"
$downloadPath = "$HOME\Downloads\crunchlabs-arduino-create-agent Setup 1.0.2.exe"
$taskName = "UnofficalCrunclabsStartupTask"
$forceReinstall = $false

# Start the script as admin if it isn't already
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-File $($MyInvocation.MyCommand.Path)"
    Exit
}

# Start script
Write-Host "Do not press buttons or switch windows as the arduino agent is setup." -ForegroundColor Green
Write-Host "Press enter if this script gets stuck" -ForegroundColor DarkGreen


# Load ability to observe window name changes to see when installer exits
Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class UserActivityHook {
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);
    }
"@

# Clear task from former runs of this script
Write-Output "Preparing for download"
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# Download if it has not been downloaded
if ($forceReinstall -or -not (Test-Path $downloadPath)) {
    # Download agent
    Write-Output "Downloading..."
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Invoke-WebRequest -Uri $crunchlabsAgentURL -OutFile $downloadPath
    Start-Sleep 1
    Write-Output "Download finished."
} else {
    Write-Output "Found agent installer already downloaded"
}

# Install if it has not already been installed
if ($forceReinstall -or -not (Test-Path $exePath)) {
    Write-Output "Starting"
    Unblock-File -Path $downloadPath -Confirm:$false # Disables run possible warnings so button presses are consistent
    Start-Sleep 1
    Start-Process $downloadPath
    Start-Sleep 3 # Wait for window to load

    # Get window, allow keypresses
    $WindowHandle = Get-Process | Where-Object { $_.MainWindowTitle -Match $WindowTitle } | Select-Object -ExpandProperty MainWindowHandle
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')

    Start-Sleep 1
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep 1
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep 5
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

    # Wait for installer to exit (as observed by window focus changing)
    $currentWindow = [UserActivityHook]::GetForegroundWindow()
    while ($true) { # repeat sleep until focus changes
        $newWindow = [UserActivityHook]::GetForegroundWindow()
        if ($newWindow -ne $currentWindow) {
            $currentWindow = $newWindow
            $windowTitle = New-Object System.Text.StringBuilder(256)
            [UserActivityHook]::GetWindowText($currentWindow, $windowTitle, $windowTitle.Capacity) | Out-Null
            Start-Sleep 2
            Write-Output "Installer is done"
            break
        }
        Start-Sleep -Seconds 1
    }

}

# Create and register task to run agent as admin
Write-Output "Configuring startup task..."
$action = New-ScheduledTaskAction -Execute $exePath # run agent
$trigger = New-ScheduledTaskTrigger -AtStartup # run at startup
$userSID = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value # run as this user's scope so the icon shows in system tray
$principal = New-ScheduledTaskPrincipal -UserId $userSID -LogonType Interactive -RunLevel Highest # but also run it as admin to avoid issues
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries:$true -DontStopIfGoingOnBatteries
$task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings 
Register-ScheduledTask -TaskName $taskName -InputObject $task | Out-Null

# And we'll also start the task now
Start-Sleep 1
Start-ScheduledTask -TaskName $taskName

# Wrap up
Write-Host "Installing finished! The crunchlabs agent should have started and be scheduled to restart on reboot." -ForegroundColor Green
Read-Host -Prompt "(press enter to exit)"
Exit

