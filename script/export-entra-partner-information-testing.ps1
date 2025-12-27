# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Nama Skrip: Get-EntraPartnerInformation_Custom
# Deskripsi: Menarik informasi partner dengan format output spesifik.
# =========================================================================

# Variabel Global dan Output
$scriptName = "EntraPartnerCustom" 
$scriptOutput = New-Object System.Collections.Generic.List[PSCustomObject]

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName


## -----------------------------------------------------------------------
## 1. PRASYARAT DAN INSTALASI MODUL
## -----------------------------------------------------------------------

Write-Host "--- 1. Memeriksa dan Menyiapkan Lingkungan PowerShell ---" -ForegroundColor Blue

# 1.1. Mengatur Execution Policy
Write-Host "1.1. Mengatur Execution Policy ke RemoteSigned..." -ForegroundColor Cyan
try {
    Set-ExecutionPolicy RemoteSigned -Scope Process -Force -ErrorAction Stop
    Write-Host " Execution Policy berhasil diatur." -ForegroundColor Green
} catch {
    Write-Error "Gagal mengatur Execution Policy: $($_.Exception.Message)"
    exit 1
}

# 1.2. Fungsi Pembantu untuk Cek dan Instal Modul
function CheckAndInstallModule {
    param([string]$ModuleName)
    Write-Host "1.$(++$global:moduleStep). Memeriksa Modul '$ModuleName'..." -ForegroundColor Cyan
    if (Get-Module -Name $ModuleName -ListAvailable) {
        Write-Host " Modul '$ModuleName' sudah terinstal." -ForegroundColor Green
    } else {
        Write-Host " Modul '$ModuleName' belum ditemukan. Menginstal..." -ForegroundColor Yellow
        Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
    }
}

$global:moduleStep = 1
CheckAndInstallModule -ModuleName "PowerShellGet"
CheckAndInstallModule -ModuleName "Microsoft.Entra"

## -----------------------------------------------------------------------
## 2. KONEKSI WAJIB (MICROSOFT ENTRA)
## -----------------------------------------------------------------------

Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Entra ---" -ForegroundColor Blue

try {
    Write-Host "Menghubungkan ke Microsoft Entra..." -ForegroundColor Yellow
    Connect-Entra -Scopes 'Organization.Read.All' -ErrorAction Stop
    Write-Host "Koneksi ke Microsoft Entra berhasil dibuat." -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung ke Microsoft Entra. $($_.Exception.Message)"
    exit 1
}

## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------

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

## -----------------------------------------------------------------------
## 4. CLEANUP, DISCONNECT, DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup, Memutus Koneksi, dan Ekspor Hasil ---" -ForegroundColor Blue

if ($scriptOutput.Count -gt 0) {
    Write-Host "Mengekspor data ke CSV..." -ForegroundColor Yellow
    try {
        # Menggunakan pemisah semicolon (;) sesuai framework sebelumnya
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Stop
        Write-Host " Data berhasil diekspor ke: $outputFilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Gagal mengekspor data ke CSV: $($_.Exception.Message)"
    }
}

Disconnect-Entra -ErrorAction SilentlyContinue
Write-Host "`nSkrip selesai dieksekusi." -ForegroundColor Yellow