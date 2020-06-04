''' Folders Handler

Private Function LPadZero(str, length)
  LPadZero = String(length - Len(str), "0") & str
End Function

Private Function DefaultFolderName()
  tm = Now()
  s = ""
  s = s & LPadZero(Year(tm), 4)
  s = s & LPadZero(Month(tm), 2)
  s = s & LPadZero(Day(tm), 2)
  s = s & "_"
  s = s & LPadZero(Hour(tm), 2)
  s = s & LPadZero(Minute(tm), 2)
  DefaultFolderName = s
End Function

' Extract arguments

Dim folderDir
folderDir = WScript.Arguments(0)

Dim folderType
folderType = WScript.Arguments(1)

' Generate the folder name

tm = Now()

tmYear = LPadZero(Year(tm), 4)
tmMon  = LPadZero(Month(tm), 2)
tmDay  = LPadZero(Day(tm), 2)

tmHour = LPadZero(Hour(tm), 2)
tmMin  = LPadZero(Minute(tm), 2)
tmSec  = LPadZero(Second(tm), 2)
  
tmDate = tmYear & tmMon & tmDay
tmTime = tmHour & tmMin & tmSec

tmSep  = "_"

folderName = ""

' Update folder name with arguments

If folderType = "--date" Then
  folderName = tmDate
ElseIf folderType = "--date-time-without-second" Then
  folderName = tmDate & tmSep & tmHour & tmMin
ElseIf folderType = "--date-time" Then
  folderName = tmDate & tmSep & tmTime
Else ' --default
  folderName = "New Folder"
End If

' Generate the final directory of the new folder

folderDir = folderDir & "\"
folderDir = folderDir & folderName

' Make directory

Set FS = CreateObject("Scripting.FileSystemObject")
If Not FS.FolderExists(folderDir) Then
  FS.CreateFolder(folderDir)
End If