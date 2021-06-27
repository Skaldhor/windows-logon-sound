# For this script NirCmd has to be installed: https://www.nirsoft.net/utils/nircmd.html

# --- Variables ---
# Path to the soundfile that should be played
$SoundFile = "C:\path\to\file.mp3"
# display name of the speakers the sound should be played on
$Speaker = "Speakers"
# Volume in Percent (0-100)
$Volume = 30
# Length of audio file in seconds (script stops 5 seconds after and kills the sound process)
$length = 5

# --- Script ---
# Setting output device and volume
$NirVolume = 65535 * $Volume * 0.01
nircmd setdefaultsounddevice $Speaker 1
Sleep 1
nircmd mutesysvolume 0 $Speaker
Sleep 1
nircmd setvolume $Speaker $NirVolume $NirVolume
Sleep 1
Write-Output "Playing... `n$SoundFile `nat device `n$Speaker `nat volume `n$Volume%"

#Play sound
Add-Type -AssemblyName presentationCore
$Player = New-Object system.windows.media.mediaplayer
$Player.open($SoundFile)
$Player.Play()

$timer = $length + 5
Sleep $timer