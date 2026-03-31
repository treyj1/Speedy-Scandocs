Set objFSO   = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

scriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
objShell.CurrentDirectory = scriptDir

reqFile = scriptDir & "\requirements.txt"
pyFile  = scriptDir & "\scandocs_tool.py"

' Install / update dependencies (hidden window, waits to finish)
objShell.Run "cmd /c python -m pip install -q -r """ & reqFile & """", 0, True

' Launch the tool (hidden window)
objShell.Run "cmd /c python """ & pyFile & """", 0, False
