'---------------------------------------------------------------
'JENKINS SOFTWARE UPDATE 
'---------------------------------------------------------------
'Uses Ninite Pro to install or update software. Implemented with Group Policy 
'to select target workstations for software updates.

Option Explicit

'Delcaring configurable variables.
Dim strNiniteLocation, strLogLocation, strNiniteFlags, blnIsSilent,_
    arrstrInclude, arrstrExclude

'---------------------------------------------------------------
'CONFIGURATION OPTIONS
'---------------------------------------------------------------
'Set the location of Ninite. It should be accessible from the PC(s) that will
'run it.
strNiniteLocation = "\\jll-dc01\staff-data\IT-Info\Login\software-updates\NiniteOne.exe"

'Set folder in which to store logs. This script will keep logs of its own and
'Ninite logs in the same file. Use GetHostHame() to return the name of the 
'PC running the script.
strLogLocation = "\\jll-dc01\staff-data\IT-Info\Login\software-updates\logs\" & GetHostName() & "\"

'Set any switch to append to the call to ninte. Do not use the silent switch.
'Use blnIsSilient variable below to run Ninite silently.
strNiniteFlags = "/updateonly /exclude KeePass"

'When set to True, Ninite will run siliently. You should not set this switch in
'strNiniteFlags. In order to keep Ninite log's (which will generate if Ninite
'runs siliently) in the same log as those of the script, the script must know
'explicitly that user desires a silient operation.
blnIsSilent = True

'If there are any elements in this array, the script will only run for those
'PCs whose hostnames are listed as elements.
'Example: arrstrInclude = Array("JIMPC", "ALICELAPTOP")
arrstrInclude = Array("ASATHER460", "CBERGSMA", "JHOHENSTEIN", "CERTEL",_ 
                      "KSNYDER-E521", "KPIECHNIK", "SWALKER", "CNELSON65",_
                      "DKINZER64", "EKRANTW7")

'The script will not run for any PCs whose hostnames are listed in this array.
arrstrExclude = Array()

'---------------------------------------------------------------
'SUBS
'---------------------------------------------------------------
Sub WriteLogFile(strMessage, strLogFilePath)
    'Records given strMessage in a given log file. Can be used with any VBS to 
    'record a single or alternating log file.
    '
    'File will be created if there is none, however partent flolder will not 
    'be. Parent folder must exist.
    '
    'Function takes a standard or network file path for strLogFilePath.
    'Examples: 
    '   C:\Folder\Subfolder\log.txt
    '   D:\Folder\Subfolder\log.txt
    '   \\network-path\subfolder\log.txt
    '   \\network-path\c$\subfolder\log.txt
    
    Dim objFso, objFile, datNow
    
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFso.OpenTextFile(strLogFilePath, 8, True)
    
    datNow = Now()
    datNow = Replace(datNow,"/","-")
   
    strMessage = datNow & "|" & GetHostName() & "|" & strMessage
    objFile.WriteLine (strMessage)
    objFile.Close
    
End Sub

Sub CreateFolder(strNewDirectory)
    'Creates a folder provided by strNewDirectory.
    'Cannot create a folder if its parent folder does not already exist.
    'Use BuildFolderHierarchy to create a folder whose parents and
    'grandparents do not exist.
    
    Dim objFSO
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CreateFolder(strNewDirectory)
    Set objFSO=nothing
    
End Sub

Sub CreateFile(strNewFile)
    'Creates a file with path provided by strNewFile.
    'Cannot create a file if its parent folder does not already exist.
    '
    'Takes a standard or network file path
    'Examples: 
    '   C:\Folder\Subfolder\
    '   D:\Folder\Subfolder
    '   \\network-path\subfolder\
    '   \\network-path\c$\subfolder\
    
    Dim objFSO, objTxtFile

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTxtFile = objFSO.CreateTextFile(strNewFile, True)
    
End Sub

'---------------------------------------------------------------
'FUNCTIONS
'---------------------------------------------------------------
Function GetHostName()
    'Returns the name of the PC.
    
    Dim wshShell
    Dim strComputerName
    
    Set wshShell = WScript.CreateObject("WScript.Shell")
    strComputerName = wshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    GetHostName = strComputerName    
    
End Function

Function DoesFolderExist(strPath)
    'Takes a directory path as strPath
    'Returns bool True if folder exists.
    'Returns bool False if folder does not exist.
    '
    'Function takes a standard or network file path
    'Examples: 
    '   C:\Folder\Subfolder\
    '   D:\Folder\Subfolder
    '   \\network-path\subfolder\
    '   \\network-path\c$\subfolder\
    
    Dim objFSO

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists(strPath) Then
        DoesFolderExist = True
    Else
        DoesFolderExist = False
    End If
End Function

Function DoesFileExist(strPath)
    'Takes a directory path as strPath
    'Returns bool True if file exists.
    'Returns bool False if file does not exist.
    '
    'Function takes a standard or network file path
    'Examples: 
    '   C:\Folder\Subfolder\file.ext
    '   D:\Folder\Subfolder\file.ext
    '   \\network-path\subfolder\file.ext
    '   \\network-path\c$\subfolder\file.ext
    
    Dim objFSO

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(strPath) Then
        DoesFileExist = True
    Else
        DoesFileExist = False
    End If  
End Function

Function BuildFolderHierarchy(strFullPath)
    'If a folder does not exist, this function will create it, and any of its
    'parents or grandparents.
    '
    'Function takes a standard or network file path
    'Examples: 
    '   C:\Folder\Subfolder\
    '   D:\Folder\Subfolder
    '   \\network-path\subfolder\
    '   \\network-path\c$\subfolder\
    '
    'Requires strFullPath to either start with a drive letter, a colon, and a 
    'backlash to indicate a local drive, or two backslashes to indicate a 
    'network location.
    '
    'Returns False on error.
    
    Dim arrstrPathParts 'Array of individual parts of strFullPath
    Dim strPart 'An individual part of the entire path
    Dim strPartial 'We'll add to this as we loop through the array
    Dim strPathType 'The type of path (network or local)
    Dim boolFirstLoop 'Whether or not we've been through 1 loop on the for stmnt

    If Len(strFullPath) = 0 Then 'Empty (but not null) value.
        BuildFolderHierarchy =  False
    ElseIf Left(strFullPath, 2) = "\\" Then
    'All network paths will start with two backslashes.
        strPathType = "network"
    ElseIf InStr(Left(strFullPath, 3), ":\") = 2 Then
    'All local paths will start with a letter and then include a colon and a 
    'backslash.
        strPathType = "local"
    Else
        BuildFolderHierarchy = False
    End If

    'This splits a file path into an array. Elements separated by the "\". 
    'Results in high-level folders at lower indexes and low-level subfolders at 
    'higher indexes.
    arrstrPathParts = split(strFullPath, "\") 

    boolFirstLoop = True
    For Each strPart In arrstrPathParts
        If len(strPart) > 0 Then
        'Excludes empty array elements produced from trailing backslash (\) 
        'at the end of strPath and the backslash that *preceeds network paths.
        'Natrually, empty elements don't count against the boolFirstLoop.
        
            If boolFirstLoop = True Then
                If strPathType = "local" Then
                'We cannot add the drive letter to strPartial in the same way 
                'we add the folders
                    strPartial = strPart
                ElseIf strPathType = "network" Then
                'Network locations need to start with two blackslashes or dir-
                'checking functions will not work.
                    strPartial = "\\" & strPart
                End If
            Else
                strPartial = strPartial & "\" & strPart
                If DoesFolderExist(strPartial) = False Then
                    CreateFolder(strPartial)
                End If
            End If
            boolFirstLoop = False
        End If
    Next
End Function

Function GetBestAvailableName(strFullPath)
    'If a file exists at strFullPath, this function will append an integer to 
    'the end of the file name to ensuer that the filename is unique. Integer 
    'will increase up to 999 until unique name is found.
    '
    'If there is no file at strFullPath, returns strFullPath.
    '
    'Function takes a standard or network file path for strFullPath.
    'Examples: 
    '   C:\Folder\Subfolder\
    '   D:\Folder\Subfolder
    '   \\network-path\subfolder\
    '   \\network-path\c$\subfolder\

    Dim i, strNewFilename
    i = 0
    strNewFilename = strFullPath
    Do While DoesFileExist(strNewFilename) = True
        strNewFilename = strFullPath & "." & i
        If i < 1000 Then
            i = i + 1
        Else
            GetBestAvailableName = False
        End If
    Loop
    GetBestAvailableName = strNewFilename

End Function

Function GetFileContents(strFullPath)
    'Opens a text file and returns the results. If there is no file at 
    'strFullPath, returns False.
    '
    'Function takes a standard or network file path for strFullPath.
    'Examples: 
    '   C:\Folder\Subfolder\
    '   D:\Folder\Subfolder
    '   \\network-path\subfolder\
    '   \\network-path\c$\subfolder\
    
    If DoesFileExist(strFullPath) Then
        Dim objFSO, strFileName, strTextLine
        Set objFSO = CreateObject("Scripting.FileSystemObject")

        'Open the file for reading.
        Set strFileName = objFSO.OpenTextFile(strFullPath, 1)

        'Loop through lines of thise file and display results.
        Do While strFileName.AtEndOfStream <> True
            strTextLine = strTextLine & vbCrLf & strFileName.ReadLine
        Loop
        strFileName.Close
        GetFileContents = strTextLine
    Else
        GetFileContents = False
    End If
End Function

Function FormatNiniteLogs(strNiniteLog)
    'Takes Ninite log file and formats it to fit in the structured document
    'that this script stores its logs in.
    
    Dim strNewNoSpaces, arrstrLogLines, blnFirstRun, strNiniteStatus,_
        strNewLogLine, i
    
    'Replace double line breaks
    strNewNoSpaces = Replace(strNiniteLog, vbCrLf & vbCrLf, vbCrLf)
    arrstrLogLines = Split(strNewNoSpaces, vbCrLf)
    
    blnFirstRun = True
    
    For Each i in arrstrLogLines
        If blnFirstRun Then 'The first element of the array is its own field.
            strNiniteStatus = "NiniteReport["
        Else
            strNewLogLine = strNewLogLine & i & ";"
        End If
        blnFirstRun = False
    Next
    FormatNiniteLogs = strNiniteStatus & strNewLogLine & "]"
End Function

Function InArray(strNeedle,arrHaystack) 
    'VBScript has no functionlaity to find an element in an array.
    'This does that.
    'Accepts needle in strNeedle, array in arrHaystack.
    'Returns position of needle in haystack.
    'Returns False if needle is not in haystack.
   
    Dim i
    For Each i In arrHaystack
        If i = strNeedle Then
            InArray = i
            Exit Function
        End If
    Next
    InArray = False
    
End Function 

'---------------------------------------------------------------
'INCLUSION LIST
'---------------------------------------------------------------
If UBound(arrstrInclude) >= 0 Then 'If arrstrInclude is not an empty array.
    'User has specified PCs to include in this run.
    'If the current PC's hostname is not in the list of approved PCs, we
    'quit the script.
    If InArray(GetHostName(),arrstrInclude) = False Then
        WScript.Quit 1
    End If
End If

'---------------------------------------------------------------
'SCRIPT LOGGING
'---------------------------------------------------------------

'Variables Declaration for log file creation....
Dim datStartTime, strLogTime, strLogFile, strScriptLog

'Footwork to create the log file.
datStartTime = Now()
strLogTime = Year(datStartTime) & "-" & _
              Month(datStartTime) & "-" & _ 
              Day(datStartTime) & "-" & _ 
              Timer

              strLogFile = "Update-" & strLogTime & ".txt"

'If the folder that we store the log file doesn't exist, create it.
If DoesFolderExist(strLogLocation) = False Then
    BuildFolderHierarchy(strLogLocation)
End If

'Set the full path for the script's log.
strScriptLog = strLogLocation & strLogFile

'Create the script's log. Function returns unique filename if selected
'filename already exists.
strScriptLog = GetBestAvailableName(strScriptLog)
If strScriptLog = False Then
    Msgbox("Script cannot create unused log filename. Clean destination "_
            & "folder: " & strLogLocation & vbCrLf & "Script will "_
            & "override log file at: " & strScriptLog)
End If
CreateFile(strScriptLog)
WriteLogFile "Process started, log file created.", strScriptLog


'---------------------------------------------------------------
'HANDLE /SILENT SWITCH
'---------------------------------------------------------------
'Checks configuration and makes log note if user configured script incorrectly.
'Updates strNiniteFlags to include /silent switch if blnIsSilent is set to
'True. Script will instruct Ninite to generate its own log file if blnIsSilent
'switch is set to True.

'Check that user requested /silent switch correctly.
If InStr(LCase(strNiniteFlags), "/silent") > 0 Then
    WriteLogFile "WARNING: Cannot accept /silent switch in strNiniteFlags " &_
                 "configuration variable. Use blnIsSilent.", _ 
                 strScriptLog
    'Ninite will likely be able to continue, so script does not close.
End If

'Set silent switch and set temp location for Ninite's log file.
If blnIsSilent Then
    
    'Create the temp Ninite log path, which only is created when Ninite
    'runs silently.
    Dim strFullNiniteLogPath
    strFullNiniteLogPath = GetBestAvailableName(strScriptLog & "-NiniteTmp")
   
   'Create the Ninite's log. Function returns unique filename if selected
   'filename already exists.
    If strFullNiniteLogPath = False Then
     
        WriteLogFile "Script cannot create unused log filename to store " &_
                      "temp Ninite logs. Clean destination folder"
                      strScriptLog
    End If    
    
    strNiniteFlags = strNiniteFlags & " /silent " & strFullNiniteLogPath
    WriteLogFile "blnIsSilient set to True.", strScriptLog
End If

WriteLogFile "strNiniteLocation: " & strNiniteLocation, strScriptLog
WriteLogFile "strNiniteFlags: " & strNiniteFlags, strScriptLog

'---------------------------------------------------------------
'RUN NINITE
'---------------------------------------------------------------
'Run Ninite at the location set in the configuration section, and with the
'switches set in configuration, including /silent switch if applicable.
'Script will instruct Ninite to generate its own log file if silent switch 
'is on.

If DoesFileExist(strNiniteLocation) Then
    Dim strStatusCode, objWshShell
    Set objWshShell = WScript.CreateObject("WScript.Shell")
    
    WriteLogFile "Set to run Ninite: " & strNiniteLocation, strScriptLog
    strStatusCode = objWshShell.Run _
                    (strNiniteLocation & " " & strNiniteFlags, 1, True)
    WriteLogFile "Ninite has finished. strStatusCode: " & strStatusCode,_
                 strScriptLog
Else
    WriteLogFile "ERROR: Cannot find Ninite at configured path: " &_
                 strNiniteLocation, strScriptLog
    Msgbox("Failed to run Ninite. Ninite is missing. Contact IT.")
End If

'---------------------------------------------------------------
'CLEAN UP NINITE'S LOG FILE, ADD TO SCRIPT'S LOG FILE
'---------------------------------------------------------------
'If we ran Ninite siliently, it created a log file. Get it and put it in our
'own logs.
If DoesFileExist(strFullNiniteLogPath) And blnIsSilent Then
'Running silient and the log files exist like they should.
    Dim strNiniteLogs
    strNiniteLogs = GetFileContents(strFullNiniteLogPath)
    WriteLogFile FormatNiniteLogs(strNiniteLogs),_
    strScriptLog
ElseIf blnIsSilent Then 
'Running silent but the Ninite log files aren't to be found.
    WriteLogFile "ERROR retrieving Ninite logs.", strScriptLog
End If


'---------------------------------------------------------------
'INSTALL LMI
'---------------------------------------------------------------
'If the LMI folder doesn't already exist, install the software.

Dim blnInstalled
blnInstalled = False

WriteLogFile "Detecting LogMeIn.", strScriptLog

If DoesFolderExist("C:\Program Files\LogMeIn") Then
    blnInstalled = True
    WriteLogFile "Detected LMI. Install canceled.", strScriptLog
End If
If DoesFolderExist("C:\Program Files (x86)\LogMeIn") Then
    blnInstalled = True
    WriteLogFile "Detected LMI. Install canceled.", strScriptLog
End If

If blnInstalled = False Then

    Dim strLMIPath, strLMICommand
    
    strLMIPath = _
    "\\jll-dc01\d$\JLL_staff\data\IT-Info\Login\software-updates\LogMeIn.msi"

    strLMICommand = _
    "msiexec.exe /i " & strLMIPath & " /quiet " &_
    "DEPLOYID=00_0boc5dynk6psnvrzkc3c9pao36hzhzu9udkb1 " &_
    "INSTALLMETHOD=5 FQDNDESC=1"
    
    WriteLogFile "LMI not detected. Installing.", strScriptLog
    WriteLogFile "strLMIPath=" & strLMIPath, strScriptLog
    WriteLogFile "strLMICommand=" & strLMICommand, strScriptLog

    Set objWshShell = WScript.CreateObject("WScript.Shell")
    strStatusCode = objWshShell.Run(strLMICommand)
    WriteLogFile "LMI Installation finished. strStatusCode:" _
                  & strStatusCode, strScriptLog

End If

WriteLogFile "Script done.", strScriptLog
                
                
