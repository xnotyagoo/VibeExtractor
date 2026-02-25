# Validacion de conexion previa
if (-not (Get-Command Get-DistributionGroup -ErrorAction SilentlyContinue)) {
    Write-Warning "No se detectan los comandos de Exchange. Por favor, conectese primero a su entorno."
    return
}

while ($true) {
    Clear-Host
    Write-Host "`n--- EXTRACTOR DE GRUPOS PROFESIONAL ---" -ForegroundColor Cyan

    # Seleccion de tipo
    do { $tipo = Read-Host "Tipo: [1] Normal / [2] Dinamico" } until ($tipo -in '1','2')

    # Busqueda y seleccion
    $term = Read-Host "Nombre o parte del grupo"
    $cmdlet = if ($tipo -eq '1') { "Get-DistributionGroup" } else { "Get-DynamicDistributionGroup" }
    $grupos = & $cmdlet -Filter "Name -like '*$term*' -or DisplayName -like '*$term*'" -ResultSize Unlimited -ErrorAction SilentlyContinue

    if (-not $grupos) { Write-Warning "No encontrado."; Start-Sleep 2; continue }

    if ($grupos.Count -gt 1) {
        Write-Host "`nResultados:" -ForegroundColor Cyan
        for ($i=0; $i -lt $grupos.Count; $i++) { Write-Host "[$($i+1)] $($grupos[$i].Name)" }
        do { $sel = Read-Host "Seleccione numero" } until ($sel -as [int] -and $sel -ge 1 -and $sel -le $grupos.Count)
        $grupo = $grupos[$sel-1]
    } else { 
        $grupo = $grupos[0]; Write-Host "Seleccionado: $($grupo.Name)" -ForegroundColor Green 
    }

    # Accion
    do { $accion = Read-Host "Accion: [1] Usuarios / [2] Codigo" } until ($accion -in '1','2')

    $nombre = $grupo.Name

    # Procesamiento
    if ($accion -eq '1') {
        # Preguntas opcionales para enriquecer el reporte
        Write-Host "`nOpciones de reporte:" -ForegroundColor Cyan
        do { $optEstado = Read-Host "Desea ver el estado de la cuenta (Activo/Deshabilitado)? [s/n]" } until ($optEstado -in 's','n')
        do { $optAtribs = Read-Host "Desea extraer Atributos Extra (Empresa, Oficina, Manager)? [s/n]" } until ($optAtribs -in 's','n')

        Write-Host "`nProcesando miembros (buscando en subgrupos si existen)..." -ForegroundColor Yellow
        
        # Expansion recursiva total
        $miembrosBrutos = @()
        $gruposRevisados = @()
        $colaGrupos = @($grupo)

        $contadorGrupos = 0
        while ($colaGrupos.Count -gt 0) {
            $grupoActual = $colaGrupos[0]
            $colaGrupos = $colaGrupos | Select-Object -Skip 1

            if ($gruposRevisados -contains $grupoActual.DistinguishedName) { continue }
            $gruposRevisados += $grupoActual.DistinguishedName
            
            $contadorGrupos++
            Write-Progress -Activity "Expandiendo estructura de grupos" -Status "Analizando: $($grupoActual.DisplayName)" -Id 1

            $miembrosTemp = @()
            if ($grupoActual.RecipientTypeDetails -match 'Dynamic') {
                $miembrosTemp = Get-Recipient -RecipientPreviewFilter $grupoActual.RecipientFilter -ResultSize Unlimited
            } else {
                $miembrosTemp = Get-DistributionGroupMember -Identity $grupoActual.Identity -ResultSize Unlimited
            }

            foreach ($m in $miembrosTemp) {
                if ($m.RecipientTypeDetails -match 'Group') {
                    $colaGrupos += $m
                } else {
                    $miembrosBrutos += $m
                }
            }
        }
        Write-Progress -Activity "Expandiendo estructura de grupos" -Completed -Id 1

        if ($miembrosBrutos) {
            Write-Host "Estructurando y limpiando datos finales..." -ForegroundColor Yellow
            
            $miembrosUnicos = $miembrosBrutos | Sort-Object PrimarySmtpAddress -Unique
            $tot = $miembrosUnicos.Count
            $c = 0

            $limpios = foreach ($usr in $miembrosUnicos) {
                $c++
                Write-Progress -Activity "Generando reporte CSV" -Status $usr.DisplayName -PercentComplete (($c/$tot)*100) -Id 2

                $obj = [ordered]@{
                    DisplayName = if ($usr.DisplayName) { $usr.DisplayName.Trim() } else { "" }
                    Department = if ($usr.Department) { $usr.Department.Trim() } else { "" }
                    Title = if ($usr.Title) { $usr.Title.Trim() } else { "" }
                    PrimarySmtpAddress = if ($usr.PrimarySmtpAddress) { $usr.PrimarySmtpAddress.ToString().Trim() } else { "" }
                }

                if ($optAtribs -eq 's') {
                    $obj.Company = if ($usr.Company) { $usr.Company.Trim() } else { "" }
                    $obj.Office = if ($usr.Office) { $usr.Office.Trim() } else { "" }
                    
                    $managerName = ""
                    if ($usr.Manager) { $managerName = ($usr.Manager -split ',')[0] -replace 'CN=','' }
                    $obj.Manager = $managerName
                }

                if ($optEstado -eq 's') {
                    $estado = "Activo"
                    if ($usr.ExchangeUserAccountControl -match 'AccountDisabled' -or $usr.RecipientTypeDetails -match 'Disabled') {
                        $estado = "Deshabilitado"
                    }
                    $obj.Estado = $estado
                }

                [PSCustomObject]$obj
            }
            Write-Progress -Activity "Generando reporte CSV" -Completed -Id 2

            $limpios = $limpios | Sort-Object DisplayName
            $ruta = "$HOME\Desktop\Miembros_$nombre.csv"
            
            $limpios | Out-GridView -Title "Resultados: $nombre (Total: $tot)"
            $limpios | Export-Csv $ruta -NoTypeInformation -Encoding UTF8
            Write-Host "-> Guardado en: $ruta" -ForegroundColor Green

        } else { Write-Warning "El grupo no contiene usuarios finales." }
        
    } else {
        $ruta = "$HOME\Desktop\Codigo_$nombre.txt"
        $codigo = if ($tipo -eq '2') { $grupo.RecipientFilter } else { "Grupo estatico. No hay codigo OPath." }
        $codigo | Out-File $ruta
        Write-Host "-> Codigo guardado en: $ruta" -ForegroundColor Green
    }

    do { $resp = Read-Host "`nDesea realizar otra consulta? [s/n]" } until ($resp -in 's','n')
    if ($resp -eq 'n') { break }
}
