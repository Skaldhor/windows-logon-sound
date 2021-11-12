# For this script NirCmd has to be installed: https://www.nirsoft.net/utils/nircmd.html
# this script uses parts of other peoples' scripts. Thanks to:
# https://superuser.com/a/704585

# --- Variables ---
# path to the soundfile that should be played
$SoundFile = "D:\path\to\Logon.mp3"
# display name of the speakers the sound should be played on
$Speaker = "Speakers"
# volume in Percent (0-100)
$Volume = 30

# --- Script ---

# define fuction for getting audio file length
function Get-AudioFileLength($path){
    $shell = New-Object -COMObject Shell.Application
    $folder = Split-Path $path
    $file = Split-Path $path -Leaf
    $shellfolder = $shell.Namespace($folder)
    $shellfile = $shellfolder.ParseName($file)
    $duration = $shellfolder.GetDetailsOf($shellfile, 27)
    $array = $duration.Split(":")
    $durationInSeconds = [int]$array[2]+([int]$array[1]*60)+([int]$array[0]*3600)
    return $durationInSeconds
}

# setting output device and volume
$length = Get-AudioFileLength $SoundFile
$NirVolume = 65535 * $Volume * 0.01
nircmd setdefaultsounddevice $Speaker 1
Sleep 1
nircmd mutesysvolume 0 $Speaker
Sleep 1
nircmd setvolume $Speaker $NirVolume $NirVolume
Sleep 1
Write-Output "Playing $SoundFile at device $Speaker at volume $Volume%."

# play sound
Add-Type -AssemblyName presentationCore
$Player = New-Object system.windows.media.mediaplayer
$Player.open($SoundFile)
$Player.Play()
Sleep ($length + 10)