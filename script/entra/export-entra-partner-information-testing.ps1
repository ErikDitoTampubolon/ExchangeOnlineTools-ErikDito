# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Nama Skrip: Get-EntraPartnerInformation_Custom
# Deskripsi: Menarik informasi partner dengan format output spesifik.
# =========================================================================

# Variabel Global dan Output
$scriptName = "EntraPartnerReport" 
$scriptOutput = New-Object System.Collections.Generic.List[PSCustomObject]

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

## ==========================================================================
##                  KONEKSI WAJIB (MICROSOFT ENTRA)
## ==========================================================================

Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Entra ---" -ForegroundColor Blue

try {
    Write-Host "Menghubungkan ke Microsoft Entra..." -ForegroundColor Yellow
    Connect-Entra -Scopes 'Organization.Read.All' -ErrorAction Stop
    Write-Host "Koneksi ke Microsoft Entra berhasil dibuat." -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung ke Microsoft Entra. $($_.Exception.Message)"
    exit 1
}

## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    Write-Host "Mengambil data Partner Information..." -ForegroundColor Cyan
    
    $partners = Get-EntraPartnerInformation -ErrorAction Stop
    
    if ($null -ne $partners) {
        foreach ($partner in $partners) {
            Write-Host "-> Menemukan Partner: $($partner.DisplayName)" -ForegroundColor Green
            
            # Format output sesuai permintaan user
            $obj = [PSCustomObject]@{
                PartnerCompanyName       = $partner.DisplayName
                companyType              = "" # Kosong sesuai contoh
                PartnerSupportTelephones = "{$(($partner.SupportTelephones -join ', '))}"
                PartnerSupportEmails     = "{$(($partner.SupportEmails -join ', '))}"
                PartnerHelpUrl           = $partner.HelpUrl
                PartnerCommerceUrl       = "" # Kosong sesuai contoh
                ObjectID                 = $partner.Id
                PartnerSupportUrl        = "" # Kosong sesuai contoh
            }
            $scriptOutput.Add($obj)
        }
    } else {
        Write-Host "Tidak ada data partner ditemukan." -ForegroundColor Yellow
    }
} catch {
    Write-Error "Terjadi kesalahan: $($_.Exception.Message)"
}

## ==========================================================================
##                          EKSPOR HASIL
## ==========================================================================

if ($scriptOutput.Count -gt 0) {
    # 1. Tentukan nama folder
    $exportFolderName = "exported_data"
    
    # 2. Ambil jalur dua tingkat di atas direktori skrip
    # Contoh: Jika skrip di C:\Users\Erik\Project\Scripts, maka ini ke C:\Users\Erik\
    $parentDir = (Get-Item $scriptDir).Parent.Parent.FullName
    
    # 3. Gabungkan menjadi jalur folder ekspor
    $exportFolderPath = Join-Path -Path $parentDir -ChildPath $exportFolderName

    # 4. Cek apakah folder 'exported_data' sudah ada di lokasi tersebut, jika belum buat baru
    if (-not (Test-Path -Path $exportFolderPath)) {
        New-Item -Path $exportFolderPath -ItemType Directory | Out-Null
        Write-Host "`nFolder '$exportFolderName' berhasil dibuat di: $parentDir" -ForegroundColor Yellow
    }

    # 5. Tentukan nama file dan jalur lengkap
    $outputFileName = "Output_$($scriptName)_$($timestamp).csv"
    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName
    
    # 6. Ekspor data
    $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    
    Write-Host "`nSemua proses selesai!" -ForegroundColor Green
    Write-Host "Laporan tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan
}