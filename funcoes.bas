Public Function PastaSistema(NumPasta As Integer) As String
    
Dim objFSO As Object
Dim objShell As Object
Dim objFolder As Object
Dim objFolderItem As Object

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace((NumPasta))
If objFolder Is Nothing Then
    PastaSistema = ""
Else
    Set objFolderItem = objFolder.Self
    PastaSistema = objFolderItem.Path
End If
End Function