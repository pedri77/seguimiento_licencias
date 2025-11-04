# Comprobar licencias

# ============================================================
# Script: Comprobar-Licencias.ps1
# Descripción: Revisa las fechas de fin en Control_Licencias.xlsx
# y muestra alertas si quedan menos de 4 meses para vencer.
# ============================================================

# Ruta del archivo Excel (ajústala según tu entorno)
$excelPath = "C:\Ruta\A\Control_Licencias.xlsx"

# Crear objeto Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)
$sheet = $workbook.Sheets.Item("Control_Licencias")

# Buscar última fila usada
$lastRow = $sheet.Cells.Find("*", $sheet.Cells.Item(1,1), $null, $null, 1, 2, $false, $false, $null).Row

# Fecha actual
$today = Get-Date
$thresholdDays = 120  # 4 meses aproximadamente

# Contador y lista de alertas
$alertas = @()

# Recorrer filas desde la segunda (evitar cabeceras)
for ($i = 2; $i -le $lastRow; $i++) {
    $producto = $sheet.Cells.Item($i, 3).Text
    $fechaFin = $sheet.Cells.Item($i, 8).Text

    if ($fechaFin -ne "") {
        try {
            $endDate = [datetime]::ParseExact($fechaFin, "dd/MM/yyyy", $null)
            $diasRestantes = ($endDate - $today).Days

            if ($diasRestantes -le $thresholdDays -and $diasRestantes -gt 0) {
                $alertas += "⚠️ $producto vence en $diasRestantes días (fecha: $fechaFin)"
            }
            elseif ($diasRestantes -le 0) {
                $alertas += "❌ $producto ya ha vencido (fecha: $fechaFin)"
            }
        }
        catch {
            Write-Host "Error al interpretar la fecha en la fila $i"
        }
    }
}

# Mostrar resultados
if ($alertas.Count -gt 0) {
    Write-Host "`n========== ALERTAS DE LICENCIAS ==========" -ForegroundColor Yellow
    $alertas | ForEach-Object { Write-Host $_ -ForegroundColor Red }
    Write-Host "==========================================`n"
} else {
    Write-Host "✅ Todas las licencias están activas." -ForegroundColor Green
}

# Cerrar Excel
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel
