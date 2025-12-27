# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Menyimpan output skrip ke file CSV dinamis di folder skrip.
# =========================================================================

# Variabel Global dan Output
$scriptName = "ExportLicenseQuotaReport" # Nama skrip yang sebenarnya
$scriptOutput = @() # Array tempat semua data hasil skrip dikumpulkan

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
# Menggunakan $PSScriptRoot memastikan file disimpan di folder yang sama dengan skrip
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"

# Penanganan kasus $PSScriptRoot tidak ada saat dijalankan dari konsol
$scriptDir = if ($PSScriptRoot) {$PSScriptRoot} else {(Get-Location).Path}
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

# ==========================================================
#                INFORMASI SCRIPT                
# ==========================================================
Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : ExportLicenseQuotaReport" -ForegroundColor Yellow
Write-Host " Field Kolom       : [LicenseName]
                     [SkuPartNumber]
                     [CapabilityStatus]
                     [TotalUnits]
                     [ConsumedUnits]
                     [AvailableUnits]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk mengambil detail semua lisensi (SKU) yang disubskripsikan dari Microsoft Graph, menghitung kuota total, jumlah yang terpakai, serta sisa lisensi yang tersedia, kemudian menampilkan hasil di konsol dan mengekspor laporan ke file CSV." -ForegroundColor Cyan
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
## 3. LOGIKA UTAMA SCRIPT ANDA DI SINI
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

# >>> GANTI BAGIAN INI DENGAN KODE UTAMA SKRIP ANDA <<<

Write-Host "3.1. Mengambil detail semua Lisensi yang Disubskripsikan (SKU)..." -ForegroundColor Cyan

try {
    # Ambil semua SKU yang disubskripsikan
    $subscribedSkus = Get-MgSubscribedSku -ErrorAction Stop

    $totalSkus = $subscribedSkus.Count
    Write-Host "Ditemukan $($totalSkus) SKU Lisensi Aktif." -ForegroundColor Green
    
    $i = 0
    foreach ($sku in $subscribedSkus) {
        $i++
        
        Write-Progress -Activity "Collecting License Quota Data" `
                       -Status "Processing License $i of ${totalSkus}: $($sku.SkuName)" `
                       -PercentComplete ([int](($i / $totalSkus) * 100))

        # Hitung Kuota
        $totalUnits = $sku.PrepaidUnits.Enabled
        $consumedUnits = $sku.ConsumedUnits
        $availableUnits = $totalUnits - $consumedUnits
        
        # Bangun objek kustom untuk diekspor
        $scriptOutput += [PSCustomObject]@{
            LicenseName = $sku.SkuName
            SkuPartNumber = $sku.SkuPartNumber
            CapabilityStatus = $sku.CapabilityStatus
            TotalUnits = $totalUnits
            ConsumedUnits = $consumedUnits
            AvailableUnits = $availableUnits
        }
    }
    
    Write-Progress -Activity "Collecting License Data Complete" -Status "Exporting Results" -Completed

    # Tampilkan di Konsol (Wajib Sesuai Permintaan)
    Write-Host "`n--- Hasil Laporan Kuota Lisensi ---" -ForegroundColor Blue
    if ($scriptOutput.Count -gt 0) {
        $scriptOutput | Format-Table -AutoSize
    } else {
        Write-Host "Tidak ada data yang tersedia untuk ditampilkan." -ForegroundColor DarkYellow
    }
    Write-Host "--------------------------------------------------------" -ForegroundColor Blue

}
catch {
    $reason = "Gagal fatal saat mengambil data Lisensi dari Microsoft Graph. Pastikan Anda memiliki scope 'Organization.Read.All' yang aktif. Error: $($_.Exception.Message)"
    Write-Error $reason
    # Tambahkan error fatal ke output jika terjadi
    $scriptOutput += [PSCustomObject]@{
        LicenseName = "FATAL ERROR"; SkuPartNumber = "N/A"; CapabilityStatus = "FAIL";
        TotalUnits = "N/A"; ConsumedUnits = "N/A"; AvailableUnits = "N/A"
    }
}

# >>> AKHIR DARI KODE UTAMA SKRIP ANDA <<<

## -----------------------------------------------------------------------
## 4. CLEANUP, DISCONNECT, DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup, Memutus Koneksi, dan Ekspor Hasil ---" -ForegroundColor Blue

# 4.1. Ekspor Hasil
# Hanya ekspor jika ada data yang valid (bukan hanya error fatal)
if ($scriptOutput.Count -gt 0 -and ($scriptOutput | Where-Object {$_.LicenseName -ne "FATAL ERROR"}).Count -gt 0) {
    Write-Host "Mengekspor $($scriptOutput.Count) baris data hasil skrip..." -ForegroundColor Yellow
    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -ErrorAction Stop
        Write-Host " Data berhasil diekspor ke:" -ForegroundColor Green
        Write-Host " $outputFilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Gagal mengekspor data ke CSV: $($_.Exception.Message)"
    }
} else {
    Write-Host "Tidak ada data lisensi valid yang dikumpulkan. Melewati ekspor." -ForegroundColor DarkYellow
}

# 4.2. Memutus koneksi Microsoft Graph
# MODIFIKASI: Menggunakan Disconnect-MgGraph
if (Get-MgContext -ErrorAction SilentlyContinue) {
    Write-Host "Memutuskan koneksi dari Microsoft Graph..." -ForegroundColor DarkYellow
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Host "Koneksi Microsoft Graph diputus." -ForegroundColor Green
}

Write-Host "`nSkrip $($scriptName) selesai dieksekusi." -ForegroundColor Yellow