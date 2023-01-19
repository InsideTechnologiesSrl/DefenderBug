<#
Author: Silvio Di Benedetto
Company: Inside Technologies
Site: www.silviodibenedetto.com
Description: this script will recreate the lost application links into Start Menu after the Defender's bug. 
this script can run with low permission because use local variable
if the application doesn't exists nothing happen
    
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

$programs = @{
    "Acrobat Reader DC" = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
    "Adobe Acrobat"            = "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    "Adobe Acrobat Distiller" = "C:\Program Files\Adobe\Acrobat DC\Acrobat\acrodist.exe"
    "Cinema 4D" = "C:\Program Files\Maxon Cinema 4D 2023\Cinema 4D.exe"
    "Blender 3.4" = ""
    "Adobe Photoshop 2023" = "C:\Program Files\Adobe\Adobe Photoshop 2023\photoshop.exe"
    "Adobe Illustrator 2023" = "C:\Program Files\Adobe\Adobe Illustrator 2023\Support Files\Contents\Windows\illustrator.exe"
    "Adobe Creative Cloud" = "C:\Program Files\Adobe\Adobe Creative Cloud\ACC\Creative Cloud.exe"
    "Firefox Private Browsing" = "C:\Program Files\Mozilla Firefox\private_browsing.exe"
    "Firefox"                  = "C:\Program Files\Mozilla Firefox\firefox.exe"
    "Google Chrome"            = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    "Microsoft Edge"           = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    "Notepad++"                = "C:\Program Files\Notepad++\notepad++.exe" #x64 version
    #"Notepad++"                = "C:\Program Files (x86)\Notepad++\notepad++.exe" #x86 version
    "Parallels Client"         = "C:\Program Files\Parallels\Client\APPServerClient.exe"
    "MobaXterm" = "C:\Program Files (x86)\Mobatek\MobaXterm\MobaXterm.exe"
    "Remote Desktop"           = "C:\Program Files\Remote Desktop\msrdcw.exe"
    "TeamViewer"               = "C:\Program Files\TeamViewer\TeamViewer.exe"
    "Royal TS6" = "C:\Program Files\Royal TS V6\royalts.exe"
    "Elgato StreamDeck" = "C:\Program Files\Elgato\StreamDeck\StreamDeck.exe"
    "Visual Studio 2022" = "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\devenv.exe"
    "Visual Studio Code" = "C:\Program Files\Microsoft VS Code\code.exe"
    "Camtasia Studio" = "C:\Program Files\TechSmith\Camtasia 2022\CamtasiaStudio.exe"
    "Camtasia 2022" = "C:\Program Files\TechSmith\Camtasia 2022\CamtasiaRecorder.exe"
    "OBS Studio" = "C:\Program Files\obs-studio\bin\64bit\obs64.exe"
    "Jabra Direct" = "C:\Program Files (x86)\Jabra\Direct6\jabra-direct.exe"
    "7-Zip File Manager" = "C:\Program Files\7-Zip\7zFM.exe"
    "Access"                   = "C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE"
    "Excel"                    = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
    "OneDrive"                 = "C:\Program Files\Microsoft OneDrive\onedrive.exe"
    "OneNote"                  = "C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE"
    "Outlook"                  = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
    "PowerPoint"               = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
    "Project"                  = "C:\Program Files\Microsoft Office\root\Office16\WINPROJ.EXE"
    "Publisher"                = "C:\Program Files\Microsoft Office\root\Office16\MSPUB.EXE"
    "Visio"                    = "C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE"
    "Word"                     = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.exe"
    "PowerShell 7 (x64)"       = "C:\Program Files\PowerShell\7\pwsh.exe"
    "SQL Server Management Studio" = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\ssms.exe"
    "Azure Data Studio" = "C:\Program Files\Azure Data Studio\azuredatastudio.exe"
    "Internet Explorer"            = "C:\Program Files (x86)\Internet Explorer\IEXPLORE.EXE"
    #"VLC Player"                   = "C:\Program Files (x86)\VideoLAN\VLC\vlc.exe" #x86 version
    "VLC Player"                   = "C:\Program Files\VideoLAN\VLC\vlc.exe" #x64 version
    "Wordpad"                      = "C:\Windows\write.EXE"
    "Relux" = "C:\Program Files (x86)\ReluxDesktop\obj\reluxPro.exe"
    "DIALux 4.13" = "C:\Program Files (x86)\DIALux\DIALux.exe"
    "DIALux evo" = "C:\Program Files\DIAL GmbH\DIALux\DIALux.exe"
    "DIALux evo x86" = "C:\Program Files\DIAL GmbH\DIALux\DIALux_x86.exe"
    "AutoCAD LT 2023" = "C:\Program Files\Autodesk\AutoCAD LT 2023\acadlt.exe"
    "Rhino 7" = "C:\Program Files\Rhino 7\System\Rhino.exe"
    "Rhino 6" = "C:\Program Files\Rhino 6\System\Rhino.exe"
    "Ultimaker Cura 5.2.1" = "C:\Program Files\Ultimaker Cura 5.2.1\Ultimaker-Cura.exe"
}

#Check for shortcuts in Start Menu, if program is available and the shortcut isn't... Then recreate the shortcut

$sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$profile_path = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -Name "ProfileImagePath").ProfileImagePath

$programs.GetEnumerator() | ForEach-Object {
    if (Test-Path -Path $_.Value) {
        if (-not (Test-Path -Path "$profile_path\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\$($_.Key).lnk")) {
            write-host ("Shortcut for {0} not found in {1}, creating it now..." -f $_.Key, $_.Value)
            $shortcut = "$profile_path\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\$($_.Key).lnk"
            $target = $_.Value
            $description = $_.Key
            $workingdirectory = (Get-ChildItem $target).DirectoryName
            $WshShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut($shortcut)
            $Shortcut.TargetPath = $target
            $Shortcut.Description = $description
            $shortcut.WorkingDirectory = $workingdirectory
            $Shortcut.Save()
            Start-Sleep -Seconds 1
        }
    }
}
