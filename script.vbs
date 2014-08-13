Set objShell = WScript.CreateObject("WScript.Shell")

Do While 1=1

 ' PART FOR TAKING SCREENSHOT

 ' "-i 0" is for Windows XP
 ' "-i 1" is for Windows 7, but somewhere works it without number
 
 cmds = objShell.Run ("c:\PSTools\psexec.exe -u DOMAIN\account-name -p password -i 0 \\computer-name c:\windows\nircmd.exe loop 1 10 savescreenshot \\placeTOsave\scr~$currdate.MM_dd_yyyy$-~$currtime.HH_mm_ss$.png /SILENT", 0)

 'waiting loop for 5 minutes
 WScript.Sleep 300000

 ' PART FOR REMOVING OLD SCREENSHOTS

 ' ################################################################
 ' # cleanup-folder.vbs
 ' #
 ' # Removes all files older than 1 week
 ' # Authored by Spencer Kuziw (s.kuziw-at-epic.ca)
 ' # Based on code by YellowShoe
 ' # Version 1.0 - Sept 23 2008
 ' ################################################################
 
 Dim fso, f, f1, fc, strComments, strScanDir
 
 ' will remove 7 days old screenshots
 strDays = 7
 
 Set fso = CreateObject("Scripting.FileSystemObject")

 Set f = fso.GetFolder("c:\placeTOsave")
 Set fc = f.Files
 For Each f1 in fc
       If DateDiff("d", f1.DateCreated, Date) > strDays Then
             fso.DeleteFile(f1)
       End If
 Next

Loop
