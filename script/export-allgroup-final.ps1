# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Nama Skrip: Get-EntraGroupsInfo
# Deskripsi: Menarik daftar semua grup dari Microsoft Entra ID.
# =========================================================================

# Variabel Global dan Output
$scriptName = "GetEntraGroups" 
$scriptOutput = New-Object System.Collections.Generic.List[PSCustomObject]

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

# ==========================================================
#                INFORMASI SCRIPT                
# ==========================================================
Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : Get-EntraGroupsInfo" -ForegroundColor Yellow
Write-Host " Field Kolom       : [GroupId]
                     [DisplayName]
                     [Description]
                     [MailEnabled]
                     [SecurityEnabled]
                     [GroupTypes]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk menarik daftar semua grup dari Microsoft Entra ID, termasuk informasi ID grup, nama tampilan, deskripsi, status mail-enabled, status security-enabled, serta tipe grup, kemudian mengekspor hasilnya ke file CSV." -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Yellow

# ==========================================================
# KONFIRMASI EKSEKUSI
# ==========================================================
$confirmation = Read-Host "Apakah Anda ingin menjalankan skrip ini? (Y/N)"

if ($confirmation -ne "Y") {
    Write-Host "`nEksekusi skrip dibatalkan oleh pengguna." -ForegroundColor Red
    return
}

## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    Write-Host "Sedang mengambil data Grup..." -ForegroundColor Cyan
    
    # Mengambil semua grup
    $groups = Get-EntraGroup -All -ErrorAction Stop
    
    if ($groups) {
        $total = $groups.Count
        Write-Host "Ditemukan $total grup." -ForegroundColor Green
        $counter = 0

        foreach ($group in $groups) {
            $counter++
            # Progres baris tunggal
            Write-Host "`r-> [$counter/$total] Memproses: $($group.DisplayName) . . ." -ForegroundColor Green -NoNewline
            
            # Membuat objek data kustom
            $obj = [PSCustomObject]@{
                GroupId          = $group.Id
                DisplayName      = $group.DisplayName
                Description      = $group.Description
                MailEnabled      = $group.MailEnabled
                SecurityEnabled  = $group.SecurityEnabled
                GroupTypes       = ($group.GroupTypes -join ", ")
            }
            $scriptOutput.Add($obj)
        }
        Write-Host "`n`nData grup berhasil dikumpulkan." -ForegroundColor Green
    } else {
        Write-Host "`nTidak ada grup yang ditemukan." -ForegroundColor Yellow
    }
} catch {
    Write-Error "Terjadi kesalahan saat mengambil data grup: $($_.Exception.Message)"
}

## -----------------------------------------------------------------------
## 4. CLEANUP, DISCONNECT, DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup, Memutus Koneksi, dan Ekspor Hasil ---" -ForegroundColor Blue

# 4.1. Ekspor Hasil
if ($scriptOutput.Count -gt 0) {
    Write-Host "Mengekspor $($scriptOutput.Count) baris data ke CSV..." -ForegroundColor Yellow
    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Stop
        Write-Host " Data berhasil diekspor ke: $outputFilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Gagal mengekspor data ke CSV: $($_.Exception.Message)"
    }
}

# 4.2. Memutus koneksi
Write-Host "Memutuskan koneksi dari Microsoft Entra..." -ForegroundColor DarkYellow
Disconnect-Entra -ErrorAction SilentlyContinue
Write-Host " Sesi Microsoft Entra diputus." -ForegroundColor Green

Write-Host "`nSkrip $($scriptName) selesai dieksekusi." -ForegroundColor Yellow