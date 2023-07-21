#"JIGGIDY"
#A simple script for e.g. keeping MS Teams availability status to show Available
#by sending the key-combo Alt+Tab (switch application window focus) every minute for a user-defined number of minutes
#@BIM4GIB 2020
##############
#
# UPDATED!!! (2023/07/21) 
# Added some sofistication: 
#   - A wee GUI for entering duration
#   - An array of commands to select from randomly, so it isn't always the same repetitive action
#   - Randomised intervals between commands
#   - A countdown timer in the console
#   - Script informs on next command interval time
#   - Added mouse jiggling action as a random command option                 
#
#Use:
# - Copy this script to a new Notepad file, save it with extension .ps1, e.g. "jiggidy.ps1"
# - Right-clik on your new file and run using Powershell

# GUI for user input
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$mins = [Microsoft.VisualBasic.Interaction]::InputBox('Enter duration (mins)', 'Duration', '60')

# Random command list
$commands = @('%{TAB}', '{NUMLOCK}', '{CAPSLOCK}', 'MoveMouse')

# Initialize random number generator
$random = New-Object -TypeName System.Random

Write-Host "Script started. Press Ctrl+C in this console window at any time to stop the script early."

for ($i = 0; $i -lt $mins; $i++) {
  # Random sleep interval between 30 and 90 seconds
  $sleepInterval = $random.Next(30, 91)

  # Print next sleep interval
  Write-Host "`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): Next command will be sent in $sleepInterval seconds"
  
  for ($j = 0; $j -lt $sleepInterval; $j++) {
    Start-Sleep -Seconds 1

    # Calculate and display time remaining every second
    $totalSecondsRemaining = ($mins - $i) * 60 - $j
    $minutesRemaining = [math]::Floor($totalSecondsRemaining / 60)
    $secondsRemaining = $totalSecondsRemaining % 60
    Write-Host "`rTime remaining: $minutesRemaining minutes, $secondsRemaining seconds                    " -NoNewline
  }

  # Random command
  $command = $commands[$random.Next($commands.Length)]

  # Execute command
  if ($command -eq 'MoveMouse') {
    # Move mouse pointer a little to the left and back
    $originalPosition = [System.Windows.Forms.Cursor]::Position
    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(($originalPosition.X-10), $originalPosition.Y)
    Start-Sleep -Milliseconds 500
    [System.Windows.Forms.Cursor]::Position = $originalPosition
    Write-Host "`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): Moved mouse"
  } else {
    # Send key command
    [System.Windows.Forms.SendKeys]::SendWait($command)
    Write-Host "`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): Sent command '$command'"
  }
}

Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): Script completed."

#Original script
#$mins = Read-Host -Prompt 'Duration (mins)? '
#$stillAlive = Add-Type -AssemblyName System.Windows.Forms
#for ($i = 0; $i -lt $mins; $i++) {
#  Start-Sleep -Seconds 60
#  $stillAlive.[System.Windows.Forms.SendKeys]::SendWait('%{TAB}')
#}
