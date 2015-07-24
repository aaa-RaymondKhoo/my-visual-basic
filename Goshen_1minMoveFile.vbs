'On Error Resume Next
Dim sNetworkShare, sLogFileName, sStr
Dim lTxtStream, oFile, oFolder, FSO
Dim destFolder, lFolder

sNetworkShare = "D:\ASOBGYN\AS-DATA\AS-Comm"

'Create FileSystemObject
Set FSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = FSO.GetFolder(sNetworkShare & "\UNETIXS")

'Open/Create Log File
If Not FSO.FolderExists(sNetworkShare & "\UNETIXS_Log") Then
     Set lFolder = FSO.CreateFolder(sNetworkShare & "\UNETIXS_Log")
End If

sLogFileName = "MoveFile Log_" & Replace(Date, "/", "-") & ".log"
sLogFileName = sNetworkShare & "\UNETIXS_Log\" & sLogFileName
Set lTxtStream = FSO.OpenTextFile(sLogFileName, 8, True)
destFolder = sNetworkShare & "\Received" 

For Each oFile In oFolder.Files

    If DateDiff("n",oFile.DateLastModified,Now) > 1 Then
        oFileName = oFile.Name
        WScript.Echo oFileName
        WScript.Echo oFolder&"\"& oFileName
        oFile.Copy (destFolder & "\" & oFile.Name)                                                        'Copy to Destination Folder
        ' Write Log
        If Not lTxtStream Is Nothing Then
            sStr = Now & " Info - " & oFile.Name & " moved to network share[" & destFolder & "]"
            lTxtStream.WriteLine (sStr)
        End If
        oFile.Delete True                             'Force delete
                    
    End If
   
Next

lTxtStream.Close

Set lTxtStream = Nothing
Set oFolder = Nothing
Set FSO = Nothing
