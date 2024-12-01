Dim i
i = 1

' Attendre 10 secondes (10 000 millisecondes)
WScript.Sleep 10000

Do While i <= 5
    MsgBox "you've been hacked"
    i = i + 1
Loop
