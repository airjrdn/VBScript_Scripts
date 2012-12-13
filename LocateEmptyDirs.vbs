Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(".")
Set sf = f.SubFolders

Dim objTextStream
Set objTextStream = fso.CreateTextFile(".\zzzDelete.bat", 2)

For Each f1 in sf
    if f1.Files.Count + f1.Subfolders.Count < 1 then
        objTextStream.WriteLine "rd " & chr(34) & f1.name & chr(34)
        s = s & f1.name & " - " & f1.Files.Count + f1.Subfolders.Count
        s = s &  vbCrLf
    end if
Next
MsgBox s
objTextStream.Close
Set objTextStream = Nothing
set fso = nothing
