# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Nama Skrip: Get-EntraDomainsReport
# Deskripsi: Mengambil daftar domain dan status verifikasinya dari Entra ID.
# =========================================================================

# Variabel Global dan Output
$scriptName = "GetEntraDomains" 
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
Write-Host " Nama Skrip        : Get-EntraDomainsReport" -ForegroundColor Yellow
Write-Host " Field Kolom       : [DomainName]
                     [IsDefault]
                     [IsVerified]
                     [Status]
                     [Authentication]
                     [IsAdminManaged]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk mengambil daftar domain dari Microsoft Entra ID beserta status verifikasi, tipe autentikasi, dan atribut pengelolaan admin, lalu mengekspor hasilnya ke file CSV." -ForegroundColor Cyan
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
## 2. KONEKSI WAJIB (MICROSOFT ENTRA)
## -----------------------------------------------------------------------

Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Entra ---" -ForegroundColor Blue

try {
    Write-Host "Menghubungkan ke Microsoft Entra. Selesaikan login pada pop-up..." -ForegroundColor Yellow
    
    # Menangani potensi konflik DLL dengan mencoba Disconnect terlebih dahulu
    Disconnect-Entra -ErrorAction SilentlyContinue
    
    # Koneksi utama
    Connect-Entra -Scopes 'Domain.Read.All' -ErrorAction Stop
    Write-Host "Koneksi ke Microsoft Entra berhasil dibuat." -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung: $($_.Exception.Message)"
    Write-Host "`nTIP: Jika error library berlanjut, tutup SEMUA jendela PowerShell lalu buka kembali." -ForegroundColor Yellow
    exit 1
}

## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    Write-Host "Mengambil data Domain..." -ForegroundColor Cyan
    $domains = Get-EntraDomain -ErrorAction Stop
    
    if ($domains) {
        $total = $domains.Count
        $counter = 0

        foreach ($domain in $domains) {
            $counter++
            Write-Host "`r-> [$counter/$total] Memproses: $($domain.Id) . . ." -ForegroundColor Green -NoNewline
            
            # Membuat objek data kustom
            $obj = [PSCustomObject]@{
                DomainName     = $domain.Id
                IsDefault      = $domain.IsDefault
                IsVerified     = $domain.IsVerified
                Status         = $domain.Status
                Authentication = $domain.AuthenticationType
                IsAdminManaged = $domain.IsAdminManaged
            }
            $scriptOutput.Add($obj)
        }
        Write-Host "`n`nData domain berhasil dikumpulkan." -ForegroundColor Green
    }
} catch {
    Write-Error "Terjadi kesalahan saat mengambil data domain: $($_.Exception.Message)"
}

## -----------------------------------------------------------------------
## 4. CLEANUP, DISCONNECT, DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup, Memutus Koneksi, dan Ekspor Hasil ---" -ForegroundColor Blue

# 4.1. Ekspor Hasil
if ($scriptOutput.Count -gt 0) {
    Write-Host "Mengekspor $($scriptOutput.Count) baris data hasil skrip..." -ForegroundColor Yellow
    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Stop
        Write-Host " Data berhasil diekspor ke: $outputFilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Gagal mengekspor data ke CSV: $($_.Exception.Message)"
    }
}

# 4.2. Memutus koneksi Entra
Write-Host "Memutuskan koneksi dari Microsoft Entra..." -ForegroundColor DarkYellow
Disconnect-Entra -ErrorAction SilentlyContinue
Write-Host " Sesi Microsoft Entra diputus." -ForegroundColor Green

Write-Host "`nSkrip $($scriptName) selesai dieksekusi." -ForegroundColor Yellow