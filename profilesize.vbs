'On Error Resume Next
Dim profilePath, fso, objFolder, profileSize
profilePath = "C:\users\rcmcda01"

Set fso = CreateObject("Scripting.FileSystemObject")
Set objFolder = fso.GetFolder(profilePath)
MsgBox ROUND(getFolderSize(objFolder)/(1024*1024),0)
' If fso.FolderExists(profilePath) Then
'     msgbox "Profile path is " & profilePath
'     Set objFolder = fso.GetFolder(profilePath)
'     profileSize = objFolder.Size
'     msgbox "Profile size is " & profileSize & " bytes"
' End If

'Function to recursive measure user profile size
Function getFolderSize(folderName)	
    On Error Resume Next
    size = 0
    hasSubfolders = False
    Set folder = fso.GetFolder(folderName)
    Err.Clear
    size = folder.Size

    If Err.Number <> 0 then   
        For Each subfolder in folder.SubFolders
            size = size + getFolderSize(subfolder.Path)
            hasSubfolders = True
        Next

        If not hasSubfolders then
            size = folder.Size
        End If
    End If

    getFolderSize = size

    Set folder = Nothing        
End Function