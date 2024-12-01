Dim objIE
Set objIE = CreateObject("InternetExplorer.Application")
objIE.Visible = True
objIE.Navigate("about:blank")

' Attendre 10 secondes (10 000 millisecondes)
WScript.Sleep 10000

' Créer le message bloquant
objIE.document.body.innerHTML = "<h1 style='color:red;'>YOU'VE BEEN HACKED</h1><p style='font-size:20px;'>Please click anywhere to close this window.</p>"

' Empêcher l'utilisateur d'interagir avec d'autres fenêtres
Do
    If objIE.Busy = False Then Exit Do
    WScript.Sleep 100
Loop

' Boucle jusqu'à ce que l'utilisateur ferme la fenêtre
Do While objIE.Visible = True
    WScript.Sleep 100
Loop

Set objIE = Nothing
