# Excel-Folder-Lister


# Explication du Script PowerShell (`folderScraper.ps1`)

Ce script PowerShell permet de scanner un répertoire (dossier) spécifié par l'utilisateur et de créer une feuille de calcul Excel contenant les noms des sous-dossiers trouvés dans ce répertoire.

## Fonctionnement
1. **Lancer le Script** : Lorsque vous exécutez le script, il vous demande d'entrer le chemin du répertoire que vous souhaitez scanner.
2. **Nom du Fichier Excel** : Ensuite, le script demande le nom sous lequel vous souhaitez enregistrer le fichier Excel.
3. **Création de la Feuille Excel** : Le script parcourt le répertoire, récupère les noms des sous-dossiers, et les inscrit dans la feuille Excel.
4. **Enregistrement du Fichier** : Le fichier Excel est ensuite sauvegardé dans le même répertoire.

# Explication du Fichier Batch (`RunScript.bat`)

Le fichier batch `RunScript.bat` est un script simple qui facilite l'exécution du script PowerShell pour les utilisateurs qui ne sont pas familiers avec l'utilisation de PowerShell.

## Fonctionnement
1. **Exécuter le Script PowerShell** : En double-cliquant sur `RunScript.bat`, le script PowerShell (`folderScraper.ps1`) est automatiquement lancé.
2. **Interaction Utilisateur** : Après l'exécution du script PowerShell, le fichier batch demande à l'utilisateur s'il souhaite quitter ou relancer le script.
   - Si l'utilisateur appuie sur 'X', le script se termine.
   - Si l'utilisateur appuie sur une autre touche, le script PowerShell est relancé.

## À Faire à l'Avenir
- **Choix de l'Emplacement de Sauvegarde** : Ajouter une option permettant à l'utilisateur de choisir l'emplacement de sauvegarde du fichier Excel, plutôt que de le sauvegarder automatiquement dans le même répertoire.
- **Inclusion des Fichier** : Ajouter une option (Flag) afin d'inclure les Fichier en plus des dossier
- **Ajout d'information supplementaire** : Ajouter des colonnes supplémentaires dans le fichier Excel pour des informations comme la taille des dossiers, la date de modification, etc.

Ces scripts rendent le processus de création de feuilles de calcul Excel à partir des noms de dossiers simple et accessible, même pour ceux qui n'ont pas d'expérience avec PowerShell.

