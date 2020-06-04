@ECHO OFF

SET PATH=%PATH%;C:\Program Files\WinRAR;C:\Program Files (x86)\WinRAR
WinRAR a -sfx -r -x".git" -z"SFX.CFG" "Folders.exe" *.*