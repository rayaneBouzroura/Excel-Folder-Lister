# script powershell qui liste un directory into an excel spreadsheet

#load l'appli excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false #pour voir ce qui se passe

#crée un nouveau workbook et une sheet
#un workbook est un fichier excel
$workbook = $excel.Workbooks.Add()
#une sheet est une feuille excel
$worksheet = $workbook.Worksheets.Item(1)

#demander a l'utilisateur de vous passer le directory (chemin du fichier example : C:\Users\user\Desktop\folder)
#tell l'user d'entrer le chemin + example
$userDirectory = Read-Host "Enter the directory path (example : C:\Users\user\Desktop\folder)"

#tell l'user d'entrer le nom du fichier excel qu'il veut creer
$excelFileName = Read-Host "Enter the excel file name (example : FolderList.xlsx)"

#ajouter xlsx a la fin du nom du fichier excel si l'user ne l'a pas fait
if( -not $excelFileName.EndsWith(".xlsx")){
    #si le nom du fichier n'a pas .xlsx a la fin, ont ajoute .xlsx a la fin
    $excelFileName = "$excelFileName.xlsx"
}


#Confirmer que le directory existe
#l'objet Test-Path permet de tester si un chemin existe
if (Test-Path $userDirectory) {
    #Recup tous les dossier dans le directory
    #l'objet Get-ChildItem permet de recuperer les items dans un directory et le flag -Directory permet de recuperer seulement les dossiers
    #si ont veut recuperer les fichiers aussi, on peut utiliser le flag -File comme
    # $folders = Get-ChildItem $userDirectory -Directory -File
    $folders = Get-ChildItem $userDirectory -Directory
    #Ecrire le nom des fichier dans la fiche excel
    $row = 1 #ont initialise la variable row a 1 (premiere ligne)
    #ont fait une boucle qui passe vers chaque item dossier dans le directory qu'ont a recuperer dans l'objet $folders
    foreach ($folder in $folders){
        #ont ecrit le nom du dossier dans la fiche excel
        #la facon dont ont ecrit dans une fiche excel est la suivante
        #ont utilise l'objet $worksheet.Cells.Item($row,1) pour specifier la cellule dans laquelle ont veut ecrire
        #row correspond a la ligne (qu'ont incremente)et 1 correspond a la colonne
        $worksheet.Cells.Item($row,1) = $folder.Name
        #ont incremente la variable row pour passer a la ligne suivante
        $row++
    }
    #sauvegarder le fichier excel dans le directory
    #objet filepath qui indique ou ont sauvegarde le fichier
    #utilisation de l'objet Join-Path qui permet de joindre deux chemins (ajoute un \ entre les deux chemins plus other truc to make it work partout)
    $excelfilePath = Join-Path $userDirectory $excelFileName
    $workbook.SaveAs($excelfilePath)

    #fermer excel
    #la methode Close permet de fermer le fichier excel et false 
    #$excel.Close($false)
    #quitter excel
    $excel.Quit()

    #techniquement ont a terminer mais maintenant ont doit ehhh..tirer la chasse sur les objets excel
    #via release com objects , sinon ils vont rester dans la memoire meme si le scripte est terminer
    #donc si ont utilise le scripte plusieurs fois, il va prendre de plus en plus de memoire et ralentir le pc et ca c'est pas cool :C
    #ont va donc liberer la memoire en utilisant la methode ReleaseComObject

    #liberer le worksheet d'abord 
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) 
    #liberer le workbook
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) 
    #liberer excel
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)

    #indiquer a l'utilisateur que le fichier excel a etait creer dans $excelfilePath
    Write-Host "le fichier excel contenent la liste des dossiers a etait creer dans $excelfilePath"
}
else {
    #si le directory n'existe pas, ont indique a l'utilisateur que le directory n'existe pas
    Write-Host "The directory $userDirectory does not exist"
}
