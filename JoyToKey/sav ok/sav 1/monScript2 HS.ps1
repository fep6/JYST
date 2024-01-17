$workbookPath = "D:\Documents\JoyToKey\mappingVirpil.xlsx"
$worksheetName = "Profil_2"

# Cr�er une instance Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Ouvrir le classeur
$workbook = $excel.Workbooks.Open($workbookPath)

# R�f�rence � la feuille de calcul
$worksheet = $workbook.Sheets | Where-Object { $_.Name -eq $worksheetName }

# ... Effectuer vos op�rations ...

# Sauvegarder et fermer le classeur
$workbook.Save()
$workbook.Close()
$excel.Quit()

# Lib�rer les objets COM
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

# D�truire les objets
Remove-Variable -Name worksheet, workbook, excel -Force
