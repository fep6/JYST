# Ajoutez cette ligne au d�but du script pour enregistrer le r�pertoire de travail
Add-Content -Path "D:\Documents\JoyToKey\log2.txt" -Value ("Working Directory: " + (Get-Location).Path)

# Importer le module ImportExcel
Import-Module "D:\Documents\JoyToKey\ImportExcel\ImportExcel.psm1" -Force

# Chemin et nom de classeur
$workbookPath = "D:\Documents\JoyToKey\mappingVirpil.xlsx"
$worksheetName = "Profil_2"

# Fonction pour r�cup�rer une instance existante d'Excel
function Get-ExcelInstance {
    try {
        [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    } catch {
        $null
    }
}

# Essayer de r�cup�rer une instance existante d'Excel
$excel = Get-ExcelInstance

# Indique si une nouvelle instance d'Excel a �t� cr��e
$newInstanceCreated = $false

# Si aucune instance n'est trouv�e, cr�er une nouvelle instance d'Excel
if ($excel -eq $null) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true  # Pour rendre Excel visible pendant le d�bogage
    $newInstanceCreated = $true
}

# Chercher le classeur dans les classeurs existants
$workbook = $excel.Workbooks | Where-Object { $_.FullName -eq $workbookPath }

if ($workbook -eq $null) {
    # Si le classeur n'est pas trouv� dans les classeurs existants, ouvrir le classeur
    $workbook = $excel.Workbooks.Open($workbookPath)
} else {
    # Chercher l'onglet existant par son nom
    $worksheet = $workbook.Sheets | Where-Object { $_.Name -eq $worksheetName }

    if ($worksheet -eq $null) {
        # Si l'onglet n'est pas trouv�, cr�er l'onglet
        $worksheet = $workbook.Sheets.Add()
        $worksheet.Name = $worksheetName
    } else {
        # Activer l'onglet existant
        $worksheet.Activate()
    }
}

# Enregistrez les modifications
$workbook.Save()

# Fermez Excel uniquement si une nouvelle instance a �t� cr��e
if ($newInstanceCreated) {
    $excel.Quit()
}

# Lib�rez les objets Excel (important pour �viter les fuites de m�moire)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

# D�truisez l'objet Excel
Remove-Variable -Name excel -Force
