Dim i
i = 1

' Attendre 10 secondes (10 000 millisecondes)
WScript.Sleep 5000

Do While i <= 5
    ' Message bloquant avec icône d'erreur et seul le bouton OK
    MsgBox "you've been hacked", vbCritical + vbOKOnly + vbSystemModal, "Warning"
    i = i + 1
Loop
