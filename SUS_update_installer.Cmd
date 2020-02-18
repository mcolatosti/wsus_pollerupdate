copy .\SUS_update.vbs c:\windows
schtasks /create /ru system /sc minute /mo 15 /tn "SUS Poller" /tr "'c:\windows\system32\cscript.exe' 'c:\windows\SUS_update.vbs'"