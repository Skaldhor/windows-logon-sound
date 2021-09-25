# this script uses parts of other peoples' scripts. Thanks to:
# https://superuser.com/a/704585
# https://stackoverflow.com/a/19348221
# https://github.com/WHumphreys/Powershell-Default-Audio-Device-Changer

# --- Variables ---
# path to the soundfile that should be played
$SoundFile = "D:\path\to\Logon.mp3"

# audio device the sound should be played on, get this ID from running the script once, then write command Get-RenderDevices
$audioDevice = "{de116135-01cd-4400-aa36-fbee7124536d}"

# volume in Percent (0-100)
$Volume = 30


# --- Script ---
# functions for setting default audio device
$cSharpSourceCode = @"
using System;
using System.Runtime.InteropServices;
public enum ERole : uint
{
    eConsole         = 0,
    eMultimedia      = 1,
    eCommunications  = 2,
    ERole_enum_count = 3
}

[Guid("F8679F50-850A-41CF-9C72-430F290290C8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IPolicyConfig
{
    // HRESULT GetMixFormat(PCWSTR, WAVEFORMATEX **);
    [PreserveSig]
    int GetMixFormat();
    
    // HRESULT STDMETHODCALLTYPE GetDeviceFormat(PCWSTR, INT, WAVEFORMATEX **);
    [PreserveSig]
	int GetDeviceFormat();
	
    // HRESULT STDMETHODCALLTYPE ResetDeviceFormat(PCWSTR);
    [PreserveSig]
    int ResetDeviceFormat();
    
    // HRESULT STDMETHODCALLTYPE SetDeviceFormat(PCWSTR, WAVEFORMATEX *, WAVEFORMATEX *);
    [PreserveSig]
    int SetDeviceFormat();
	
    // HRESULT STDMETHODCALLTYPE GetProcessingPeriod(PCWSTR, INT, PINT64, PINT64);
    [PreserveSig]
    int GetProcessingPeriod();
	
    // HRESULT STDMETHODCALLTYPE SetProcessingPeriod(PCWSTR, PINT64);
    [PreserveSig]
    int SetProcessingPeriod();
	
    // HRESULT STDMETHODCALLTYPE GetShareMode(PCWSTR, struct DeviceShareMode *);
    [PreserveSig]
    int GetShareMode();
	
    // HRESULT STDMETHODCALLTYPE SetShareMode(PCWSTR, struct DeviceShareMode *);
    [PreserveSig]
    int SetShareMode();
	 
    // HRESULT STDMETHODCALLTYPE GetPropertyValue(PCWSTR, const PROPERTYKEY &, PROPVARIANT *);
    [PreserveSig]
    int GetPropertyValue();
	
    // HRESULT STDMETHODCALLTYPE SetPropertyValue(PCWSTR, const PROPERTYKEY &, PROPVARIANT *);
    [PreserveSig]
    int SetPropertyValue();
	
    // HRESULT STDMETHODCALLTYPE SetDefaultEndpoint(__in PCWSTR wszDeviceId, __in ERole role);
    [PreserveSig]
    int SetDefaultEndpoint(
        [In] [MarshalAs(UnmanagedType.LPWStr)] string wszDeviceId, 
        [In] [MarshalAs(UnmanagedType.U4)] ERole role);
	
    // HRESULT STDMETHODCALLTYPE SetEndpointVisibility(PCWSTR, INT);
    [PreserveSig]
	int SetEndpointVisibility();
}
[ComImport, Guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9")]
internal class _CPolicyConfigClient
{
}
public class PolicyConfigClient
{
    public static int SetDefaultDevice(string deviceID)
    {
        IPolicyConfig _policyConfigClient = (new _CPolicyConfigClient() as IPolicyConfig);
	try
        {
            Marshal.ThrowExceptionForHR(_policyConfigClient.SetDefaultEndpoint(deviceID, ERole.eConsole));
		    Marshal.ThrowExceptionForHR(_policyConfigClient.SetDefaultEndpoint(deviceID, ERole.eMultimedia));
		    Marshal.ThrowExceptionForHR(_policyConfigClient.SetDefaultEndpoint(deviceID, ERole.eCommunications));
		    return 0;
        }
        catch
        {
            return 1;
        }
    }
}
"@

add-type -TypeDefinition $cSharpSourceCode

function Get-CaptureDevices 
{
    Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\*\Properties\" | 
    Select @{Name="CaptureDeviceId";Expression={$_.PSParentPath.substring($_.PSParentPath.length - 38, 38)}}, 
           @{Name="CaptureDeviceName";Expression={$_."{a45c254e-df1c-4efd-8020-67d146a850e0},2"}}, 
           @{Name="CaptureDeviceInterface";Expression={$_."{b3f8fa53-0004-438e-9003-51a46e139bfc},6"}} | 
    Format-List
}

function Get-CaptureDeviceId 
{
    Param
    (
        [parameter(Mandatory=$true)]
        [string[]]
        $captureDeviceName,
        [parameter(Mandatory=$true)]
        [string[]]
        $captureDeviceInterface
    )

    Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\capture\*\Properties\" |
    Where {($_."{a45c254e-df1c-4efd-8020-67d146a850e0},2" -eq $captureDeviceName) -and ($_."{b3f8fa53-0004-438e-9003-51a46e139bfc},6" -eq $captureDeviceInterface)} |
    ForEach {$_.PSParentPath.substring($_.PSParentPath.length - 38, 38)}
}

function Get-RenderDevices 
{
    Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\*\Properties\" | 
    Select @{Name="CaptureDeviceId";Expression={$_.PSParentPath.substring($_.PSParentPath.length - 38, 38)}}, 
           @{Name="CaptureDeviceName";Expression={$_."{a45c254e-df1c-4efd-8020-67d146a850e0},2"}}, 
           @{Name="CaptureDeviceInterface";Expression={$_."{b3f8fa53-0004-438e-9003-51a46e139bfc},6"}} | 
    Format-List
}

function Get-RenderDeviceId 
{
    Param
    (
        [parameter(Mandatory=$true)]
        [string[]]
        $renderDeviceName,
        [parameter(Mandatory=$true)]
        [string[]]
        $renderDeviceInterface
    )

    Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\*\Properties\" |
    Where {($_."{a45c254e-df1c-4efd-8020-67d146a850e0},2" -eq $renderDeviceName) -and ($_."{b3f8fa53-0004-438e-9003-51a46e139bfc},6" -eq $renderDeviceInterface)} |
    ForEach {$_.PSParentPath.substring($_.PSParentPath.length - 38, 38)}
}

function Set-DefaultAudioDevice
{
    Param
    (
        [parameter(Mandatory=$true)]
        [string[]]
        $deviceId
    )

    If ([PolicyConfigClient]::SetDefaultDevice("{0.0.0.00000000}.$deviceId") -eq 0)
    {
        Write-Host "SUCCESS: The default audio device has been set."
    }
    Else
    {
        Write-Host "ERROR: There has been a problem setting the default audio device."
    }
}


# functions for setting volume and mute

Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;

[Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioEndpointVolume {
  // f(), g(), ... are unused COM method slots. Define these if you care
  int f(); int g(); int h(); int i();
  int SetMasterVolumeLevelScalar(float fLevel, System.Guid pguidEventContext);
  int j();
  int GetMasterVolumeLevelScalar(out float pfLevel);
  int k(); int l(); int m(); int n();
  int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, System.Guid pguidEventContext);
  int GetMute(out bool pbMute);
}
[Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice {
  int Activate(ref System.Guid id, int clsCtx, int activationParams, out IAudioEndpointVolume aev);
}
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator {
  int f(); // Unused
  int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
}
[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }

public class Audio {
  static IAudioEndpointVolume Vol() {
    var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
    IMMDevice dev = null;
    Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(/*eRender*/ 0, /*eMultimedia*/ 1, out dev));
    IAudioEndpointVolume epv = null;
    var epvid = typeof(IAudioEndpointVolume).GUID;
    Marshal.ThrowExceptionForHR(dev.Activate(ref epvid, /*CLSCTX_ALL*/ 23, 0, out epv));
    return epv;
  }
  public static float Volume {
    get {float v = -1; Marshal.ThrowExceptionForHR(Vol().GetMasterVolumeLevelScalar(out v)); return v;}
    set {Marshal.ThrowExceptionForHR(Vol().SetMasterVolumeLevelScalar(value, System.Guid.Empty));}
  }
  public static bool Mute {
    get { bool mute; Marshal.ThrowExceptionForHR(Vol().GetMute(out mute)); return mute; }
    set { Marshal.ThrowExceptionForHR(Vol().SetMute(value, System.Guid.Empty)); }
  }
}
'@


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


# getting length of audio file in seconds
$length = Get-AudioFileLength $SoundFile

# set default audio device
Set-DefaultAudioDevice $audioDevice

# unmute audio device
[Audio]::Mute = $false

# set volume
[Audio]::Volume = ($Volume * 0.01)

# output summary
Write-Output "Playing $SoundFile at volume $Volume%."

# play sound
Add-Type -AssemblyName presentationCore
$Player = New-Object system.windows.media.mediaplayer
$Player.open($SoundFile)
$Player.Play()
Sleep ($length + 10)