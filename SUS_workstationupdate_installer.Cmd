copy .\SUS_workstationupdate.vbs c:\windows
schtasks /create /ru system /sc minute /mo 30 /tn "SUS Poller" /tr "'c:\windows\system32\cscript.exe' 'c:\windows\SUS_workstationupdate.vbs'"