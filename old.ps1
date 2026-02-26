    while ($true) {
        Clear-Host
        Write-Host "`n--- Extractor de grupos de exchange ---" -ForegroundColor Cyan

        # 1. Seleccion de tipo
        do { $tipo = Read-Host "Tipo: [1] Normal / [2] Dinamico" } until ($tipo -in '1','2')

        # 2. Busqueda y seleccion
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
            $grupo = $grupos[0]; Write-Host "Seleccionado automaticamente: $($grupo.Name)" -ForegroundColor Green 
        }

        # 3. Accion
        do { $accion = Read-Host "Accion: [1] Usuarios / [2] Codigo" } until ($accion -in '1','2')

        # 4. Procesamiento
        $nombre = $grupo.Name
        if ($accion -eq '1') {
            Write-Host "Procesando miembros..." -ForegroundColor Yellow
            
            if ($tipo -eq '2') {
                $miembros = Get-Recipient -RecipientPreviewFilter $grupo.RecipientFilter -ResultSize Unlimited
            } else {
                do { $hib = Read-Host "Listas dinamicas anidadas? [s/n]" } until ($hib -in 's','n')
                $brutos = Get-DistributionGroupMember -Identity $nombre -ResultSize Unlimited
                
                if ($hib -eq 's') {
                    $tot = $brutos.Count; $c = 0
                    $miembros = foreach ($it in $brutos) {
                        $c++; Write-Progress -Activity "Procesando" -Status $it.DisplayName -PercentComplete (($c/$tot)*100)
                        if ($it.RecipientTypeDetails -match 'Dynamic') {
                            Get-Recipient -RecipientPreviewFilter (Get-DynamicDistributionGroup $it.Name).RecipientFilter -ResultSize Unlimited
                        } else { $it }
                    }
                    Write-Progress -Activity "Procesando" -Completed
                } else { $miembros = $brutos }
            }

            if ($miembros) {
                $ruta = "$HOME\Desktop\Miembros_$nombre.csv"
                # Limpieza y formateo comprimido en un solo bloque
                $limpios = $miembros | Select-Object @{n='DisplayName';e={$_.DisplayName.Trim()}}, 
                                                     @{n='Department';e={$_.Department.Trim()}}, 
                                                     @{n='Title';e={$_.Title.Trim()}}, 
                                                     @{n='PrimarySmtpAddress';e={$_.PrimarySmtpAddress.ToString().Trim()}} | Sort-Object DisplayName -Unique
                
                $limpios | Out-GridView -Title $nombre
                $limpios | Export-Csv $ruta -NoTypeInformation -Encoding UTF8
                Write-Host "-> Guardado en: $ruta" -ForegroundColor Green
            } else { Write-Warning "Grupo vacio." }
        } else {
            $ruta = "$HOME\Desktop\Codigo_$nombre.txt"
            $codigo = if ($tipo -eq '2') { $grupo.RecipientFilter } else { "Grupo estatico. No hay codigo OPath." }
            $codigo | Out-File $ruta
            Write-Host "-> Guardado en: $ruta" -ForegroundColor Green
        }

        do { $resp = Read-Host "`nOtro? [s/n]" } until ($resp -in 's','n')
        if ($resp -eq 'n') { break }
    }
}
