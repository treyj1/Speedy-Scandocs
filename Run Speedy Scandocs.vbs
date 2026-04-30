Set objFSO   = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

scriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
objShell.CurrentDirectory = scriptDir

reqFile    = scriptDir & "\requirements.txt"
pyFile     = scriptDir & "\scandocs_tool.py"
markerFile = scriptDir & "\.deps_installed"

' Only run pip install on first launch, or when requirements.txt has changed
' since the last successful install. Keeps normal launches instant.
needInstall = True
If objFSO.FileExists(markerFile) And objFSO.FileExists(reqFile) Then
    reqMTime    = objFSO.GetFile(reqFile).DateLastModified
    markerMTime = objFSO.GetFile(markerFile).DateLastModified
    If markerMTime >= reqMTime Then
        needInstall = False
    End If
End If

If needInstall Then
    objShell.Run "cmd /c python -m pip install -q -r """ & reqFile & """", 0, True
    ' Mark install as successful so subsequent launches skip pip
    Set f = objFSO.CreateTextFile(markerFile, True)
    f.WriteLine "installed"
    f.Close
End If

' Launch the tool (hidden window)
objShell.Run "cmd /c python """ & pyFile & """", 0, False
