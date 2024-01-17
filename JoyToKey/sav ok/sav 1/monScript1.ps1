# Ajoutez cette ligne au début du script pour enregistrer le répertoire de travail
Add-Content -Path "D:\Documents\JoyToKey\log1.txt" -Value ("Working Directory: " + (Get-Location).Path)

# Importer le module ImportExcel
Import-Module "D:\Documents\JoyToKey\ImportExcel\ImportExcel.psm1" -Force

# Chemin et nom de classeur
$workbookPath = "D:\Documents\JoyToKey\mappingVirpil.xlsx"
$worksheetName = "Profil_1"

# Fonction pour récupérer une instance existante d'Excel
function Get-ExcelInstance {
    try {
        [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    } catch {
        $null
    }
}

# Essayer de récupérer une instance existante d'Excel
$excel = Get-ExcelInstance

# Indique si une nouvelle instance d'Excel a été créée
$newInstanceCreated = $false

# Si aucune instance n'est trouvée, créer une nouvelle instance d'Excel
if ($excel -eq $null) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true  # Pour rendre Excel visible pendant le débogage
    $newInstanceCreated = $true
}

# Chercher le classeur dans les classeurs existants
$workbook = $excel.Workbooks | Where-Object { $_.FullName -eq $workbookPath }

if ($workbook -eq $null) {
    # Si le classeur n'est pas trouvé dans les classeurs existants, ouvrir le classeur
    $workbook = $excel.Workbooks.Open($workbookPath)
} else {
    # Chercher l'onglet existant par son nom
    $worksheet = $workbook.Sheets | Where-Object { $_.Name -eq $worksheetName }

    if ($worksheet -eq $null) {
        # Si l'onglet n'est pas trouvé, créer l'onglet
        $worksheet = $workbook.Sheets.Add()
        $worksheet.Name = $worksheetName
    } else {
        # Activer l'onglet existant
        $worksheet.Activate()
    }
}

# Enregistrez les modifications
$workbook.Save()

# Fermez Excel uniquement si une nouvelle instance a été créée
if ($newInstanceCreated) {
    $excel.Quit()
}

# Libérez les objets Excel (important pour éviter les fuites de mémoire)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

# Détruisez l'objet Excel
Remove-Variable -Name excel -Force
