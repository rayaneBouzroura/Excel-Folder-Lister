@ECHO off
:: ce script batch permet de lancer le script ps folderScraper 
:ScriptStart
:: Lancement du script ps
:: %~dp0 donne le chemin du dossier où se trouve le fichier batch
Powershell -NoProfile -ExecutionPolicy Bypass -File  "%~dp0folderScraper.ps1"
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