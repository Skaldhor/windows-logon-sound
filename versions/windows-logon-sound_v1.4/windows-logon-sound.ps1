# this script requires you to install the following module:
# https://www.powershellgallery.com/packages/AudioDeviceCmdlets/3.0.0.4

# --- Variables ---
# path to the soundfile that should be played
$SoundFilePath = "D:\path\to\Logon.mp3"

# audio device the sound should be played on, get this ID from running the script once, then write command Get-RenderDevices
# run 'Get-AudioDevice -List' in order to find out your index
$AudioDeviceIndex = 1

# volume in Percent (0-100)
$Volume = 30


# --- Script ---
# helpfer function to define fuction for getting audio file length
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
# getting length of audio file in seconds
$Length = Get-AudioFileLength $SoundFilePath

# import AudioDeviceCmdlets module
try{
    Import-Module -Name AudioDeviceCmdlets -ErrorAction Stop
}catch{
    Write-Host "Can't import module ""AudioDeviceCmdlets"". Is it installed?"
    Write-Host "If not, istall it with 'Install-Module -Name AudioDeviceCmdlets'."
}

# get current playback device info
$CurrentDevice = (Get-AudioDevice -Playback).Index
$CurrentMute = Get-AudioDevice -PlaybackMute
[String]$CurrentVolume = Get-AudioDevice -PlaybackVolume

# set temporary plaback device
Set-AudioDevice -Index $AudioDeviceIndex
Set-AudioDevice -PlaybackMute $false
Set-AudioDevice -PlaybackVolume $Volume

# output summary
Write-Output "Playing $SoundFilePath at volume $Volume%."

# play sound
Add-Type -AssemblyName presentationCore
$Player = New-Object system.windows.media.mediaplayer
$Player.Open($SoundFilePath)
$Player.Play()
Start-Sleep ($Length + 10)

# restore original state
Set-AudioDevice -Index $CurrentDevice
Set-AudioDevice -PlaybackMute $CurrentMute
[Int32]$CurrentVolume = $CurrentVolume.Split("%")[0]
Set-AudioDevice -PlaybackVolume $CurrentVolume
