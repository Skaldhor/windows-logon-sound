# Windows-logon-sound
This script can play a soundfile of your choice on the speakers of your choice at the volume of your choice and can be executed by task scheduler.
## Installation
1. download the .ps1 script
2. edit it, change the variables to your desired settings (path to soundfile, speakers the sound should be played on, volume in percent) and save it
3. create a task in task scheduler that runs the script whenever you start the PC and sign in
## Credits
- Changing volume with powershell: https://stackoverflow.com/a/19348221
- Set audio device with powershell: https://github.com/WHumphreys/Powershell-Default-Audio-Device-Changer
- Get audio file length with powershell: https://superuser.com/a/704585