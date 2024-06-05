# This file downloads Main.ps1 to disk (which is necessary to run itself as admin)
#
# Quick run command from Run dialouge: `powershell -W H -Ex B -c "iex (iwr 'https://raw.githubusercontent.com/Reginald-Gillespie/UCAAIW/main/runner.ps1').Content`
#

# Save to disk
$scriptURL = "https://raw.githubusercontent.com/Reginald-Gillespie/UCAAIW/main/Main.ps1"
$saveLoc = "C:\Temp\Main.ps1"
Invoke-WebRequest -Uri $scriptURL -OutFile $saveLoc

# Start
Invoke-Expression -Command $saveLoc

