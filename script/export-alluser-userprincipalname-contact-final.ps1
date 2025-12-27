# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.5)
# Deskripsi: Mengambil DisplayName, UPN, dan Kontak dari SEMUA User Aktif.
# =========================================================================

# Variabel Global dan Output
$scriptName = "AllActiveUsersContactReport" 
$scriptOutput = @() 
$global:moduleStep = 1

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = if ($PSScriptRoot) {$PSScriptRoot} else {(Get-Location).Path}
$outputFileName = "Output_${scriptName}_${timestamp}.csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

# ==========================================================
#                INFORMASI SCRIPT                
# ==========================================================
Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : AllActiveUsersContactReport" -ForegroundColor Yellow
Write-Host " Field Kolom       : [DisplayName]
                     [UserPrincipalName]
                     [Contact]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk mengambil data seluruh pengguna aktif, termasuk DisplayName, UserPrincipalName, dan informasi kontak (BusinessPhones serta MobilePhone). Hasil laporan ditampilkan di konsol dengan progres bar dan diekspor otomatis ke file CSV." -ForegroundColor Cyan
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
## 2. PRASYARAT DAN KONEKSI
## -----------------------------------------------------------------------

Write-Host "--- 1. Menyiapkan Lingkungan ---" -ForegroundColor Blue 
if (-not (Get-MgContext -ErrorAction SilentlyContinue)) { 
    Connect-MgGraph -Scopes "User.Read.All" -ErrorAction Stop | Out-Null
}


## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------
Write-Host "`n--- 3. Memulai Logika Utama Skrip: ${scriptName} ---" -ForegroundColor Magenta 

try {
    Write-Host "Mengambil data seluruh pengguna aktif..." -ForegroundColor Cyan
    
    $selectProperties = @("UserPrincipalName", "DisplayName", "BusinessPhones", "MobilePhone", "AccountEnabled")

    # Ambil semua user aktif
    $allUsers = Get-MgUser -Filter "accountEnabled eq true" -All -Property $selectProperties -ErrorAction Stop
    
    $totalUsers = $allUsers.Count
    Write-Host "Berhasil menemukan ${totalUsers} pengguna aktif." -ForegroundColor Green
    
    $i = 0
    foreach ($user in $allUsers) {
        $i++
        Write-Progress -Activity "Memproses Data Kontak" -Status "User: $($user.UserPrincipalName)" -PercentComplete ([int](($i / $totalUsers) * 100))

        # Gabungkan Nomor Telepon
        $phones = @()
        if ($user.BusinessPhones) { $phones += ($user.BusinessPhones -join ", ") }
        if ($user.MobilePhone) { $phones += $user.MobilePhone }
        
        $contactInfo = if ($phones.Count -gt 0) { $phones -join " | " } else { "-" }

        # MODIFIKASI: Urutan kolom sesuai permintaan: DisplayName, UPN, Contact
        $scriptOutput += [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Contact           = $contactInfo
        }
    }
}
catch {
    Write-Error "Kesalahan saat pengambilan data: $($_.Exception.Message)"
}

## -----------------------------------------------------------------------
## 4. EKSPOR HASIL
## -----------------------------------------------------------------------
Write-Host "`n--- 4. Cleanup dan Ekspor Hasil ---" -ForegroundColor Blue 

if ($scriptOutput.Count -gt 0) { 
    Write-Host "Mengekspor $($scriptOutput.Count) baris data..." -ForegroundColor Yellow 
    try { 
        # Delimiter titik koma (;) agar otomatis rapi saat dibuka di Excel Indonesia
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Stop 
        Write-Host " Laporan berhasil disimpan di:" -ForegroundColor Green 
        Write-Host " $outputFilePath" -ForegroundColor Cyan 
    } 
    catch { 
        Write-Error "Gagal ekspor CSV: $($_.Exception.Message)" 
    } 
}

# Sesi tetap dibuka untuk integrasi menu utama Anda
Write-Host "`nSkrip ${scriptName} selesai dieksekusi." -ForegroundColor Yellow