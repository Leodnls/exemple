Dim i
i = 1

' Attendre 10 secondes (10 000 millisecondes)
WScript.Sleep 10000

Do While i <= 5
    MsgBox "YOU'VE BEEN HACKED", vbCritical + vbOKOnly + vbSystemModal, "Warning"
    i = i + 1
Loop
