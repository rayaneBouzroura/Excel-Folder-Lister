@ECHO off
:: ce script batch permet de lancer le script ps folderScraper 
:ScriptStart
:: Lancement du script ps
Powershell -NoProfile -ExecutionPolicy Bypass -File  "E:\scripts ps\folder excel ps\folderScraper.ps1"
::echo vide pour sauter une ligne
ECHO.
ECHO Appuer sur 'X' puis entrer pour quitter , appuyer sur n'importe quelle touche puis entrer pour relancer le script :D 

:: /P pour demander une entreer d'une seul lettre a l'utilisateur 
SET /P UserChoice=Choice

:: Check if the user input is not 'X' and restart if it is not
IF /I NOT "%UserChoice%"=="X" GOTO :ScriptStart

:: Exit the script if 'X' was pressed
ECHO Exiting script...
:: A timeout to give the user a chance to see the exit message
TIMEOUT /T 3