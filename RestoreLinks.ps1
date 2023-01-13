#Windows Powershell 7
if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Windows PowerShell 7.lnk")){  
 $ComObj = New-Object -ComObject WScript.Shell
    $ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Windows PowerShell 7.lnk")
    $ShortCut.TargetPath = "C:\Program Files\PowerShell\7\pwsh.exe"
    $ShortCut.Description = "Windows PowerShell 7"
    $ShortCut.FullName 
    $ShortCut.WindowStyle = 1
    $ShortCut.Save()
    }
#Edge
if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Microsoft Edge.lnk")){  
 $ComObj = New-Object -ComObject WScript.Shell
    $ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Microsoft Edge.lnk")
    $ShortCut.TargetPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    $ShortCut.Description = "Edge"
    $ShortCut.FullName 
    $ShortCut.WindowStyle = 1
    $ShortCut.Save()
    }
#OneDrive
if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\OneDrive.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\OneDrive.lnk")
	$ShortCut.TargetPath = "C:\Program Files\Microsoft OneDrive\onedrive.exe"
	$ShortCut.Description = "OneDrive"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#OneNote
if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\OneNote.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\OneNote.lnk")
	$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\OneNote.exe"
	$ShortCut.Description = "OneNote"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#PowerPoint	
if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Powerpoint.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Powerpoint.lnk")
	$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\PowerPNT.exe"
	$ShortCut.Description = "PowerPoint"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#Word
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Word.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Word.lnk")
	$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\Winword.EXE"
	$ShortCut.Description = "Word"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#Outlook
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Outlook.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Outlook.lnk")
	$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\Outlook.exe"
	$ShortCut.Description = "Outlook"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#Excel
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Excel.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Excel.lnk")
	$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
	$ShortCut.Description = "Excel"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#Parallels RAS
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Parallels RAS.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Parallels RAS.lnk")
	$ShortCut.TargetPath = "C:\Program Files\Parallels\Client\APPServerClient.exe"
	$ShortCut.Description = "Parallels RAS"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#SQL Management Studio
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\SQL Server Management Studio.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\SQL Server Management Studio.lnk")
	$ShortCut.TargetPath = "C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\ssms.exe"
	$ShortCut.Description = "SQL Server Management Studio"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#Azure Data Studio
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Azure Data Studio.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Azure Data Studio.lnk")
	$ShortCut.TargetPath = "C:\Program Files\Azure Data Studio\azuredatastudio.exe"
	$ShortCut.Description = "Azure Data Studio"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#Adobe Photoshop 2023
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Adobe Photoshop 2023.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Adobe Photoshop 2023.lnk")
	$ShortCut.TargetPath = "C:\Program Files\Adobe\Adobe Photoshop 2023\photoshop.exe"
	$ShortCut.Description = "Adobe Photoshop 2023"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#Adobe Illustrator 2023
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Adobe Illustrator 2023.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Adobe Illustrator 2023.lnk")
	$ShortCut.TargetPath = "C:\Program Files\Adobe\Adobe Illustrator 2023\Support Files\Contents\Windows\illustrator.exe"
	$ShortCut.Description = "Adobe Illustrator 2023"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#Adobe DC 2023
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Adobe DC.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Adobe DC.lnk")
	$ShortCut.TargetPath = "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
	$ShortCut.Description = "Adobe DC"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#RoyalTS 6
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Royal TS6.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Royal TS6.lnk")
	$ShortCut.TargetPath = "C:\Program Files\Royal TS V6\royalts.exe"
	$ShortCut.Description = "Royal TS6"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#Elgato Streamdeck
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Elgato StreamDeck.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Elgato StreamDeck.lnk")
	$ShortCut.TargetPath = "C:\Program Files\Elgato\StreamDeck\StreamDeck.exe"
	$ShortCut.Description = "Elgato StreamDeck"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#Visual Studio 2022
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Visual Studio 2022.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Visual Studio 2022.lnk")
	$ShortCut.TargetPath = "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\IDE\devenv.exe"
	$ShortCut.Description = "Visual Studio 2022"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#Visual Studio Code
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Camtasia Studio.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Camtasia Studio.lnk")
	$ShortCut.TargetPath = "C:\Program Files\TechSmith\Camtasia 2022\CamtasiaStudio.exe"
	$ShortCut.Description = "Camtasia Studio"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#Camtasia Studio
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Camtasia Recorder.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Camtasia Recorder.lnk")
	$ShortCut.TargetPath = "C:\Program Files\TechSmith\Camtasia 2022\CamtasiaRecorder.exe"
	$ShortCut.Description = "Camtasia Recorder"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#Jabra Direct
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Jabra Direct.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Jabra Direct.lnk")
	$ShortCut.TargetPath = "C:\Program Files (x86)\Jabra\Direct6\jabra-direct.exe"
	$ShortCut.Description = "Jabra Direct"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
#Notepad++
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\Notepad++.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\Notepad++.lnk")
	$ShortCut.TargetPath = "C:\Program Files\Notepad++\notepad++.exe"
	$ShortCut.Description = "Notepad++"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
# 7-Zip
	if(!(Test-Path -Path $env:USERPROFILE+"\Start Menu\Programs\7-Zip File Manager.lnk")){  
	$ComObj = New-Object -ComObject WScript.Shell
	$ShortCut = $ComObj.CreateShortcut($env:USERPROFILE + "\Start Menu\Programs\7-Zip File Manager.lnk")
	$ShortCut.TargetPath = "C:\Program Files\7-Zip\7zFM.exe"
	$ShortCut.Description = "7-Zip File Manager"
	$ShortCut.FullName 
	$ShortCut.WindowStyle = 1
	$ShortCut.Save()
	}
