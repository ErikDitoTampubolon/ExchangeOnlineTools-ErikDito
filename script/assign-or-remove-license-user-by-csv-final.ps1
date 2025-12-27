# =========================================================================
# LISENSI MICROSOFT GRAPH ASSIGNMENT/REMOVAL SCRIPT V19.6
# AUTHOR: Erik Dito Tampubolon - TelkomSigma
# =========================================================================

# 1. Konfigurasi File Input & Path
$inputFileName = "UserPrincipalName.csv"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$inputFilePath = Join-Path -Path $scriptDir -ChildPath $inputFileName
$defaultUsageLocation = 'ID'


# ==========================================================
#                INFORMASI SCRIPT                
# ==========================================================
Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : Assign or Remove License User" -ForegroundColor Yellow
Write-Host " Field Kolom       : [UserPrincipalName]
                     [DisplayName]
                     [Status]
                     [Reason]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk melakukan otomatisasi proses pemberian (assign) atau penghapusan (remove) lisensi menggunakan daftar CSV." -ForegroundColor Cyan
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
## 2. KONEKSI KE MICROSOFT GRAPH
## -----------------------------------------------------------------------
$requiredScopes = "User.ReadWrite.All", "Organization.Read.All"
Write-Host "`n--- 1. Memeriksa Koneksi ---" -ForegroundColor Blue

if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
    Write-Host "Menghubungkan ke Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes $requiredScopes -ContextScope Process -ErrorAction Stop | Out-Null
}
Write-Host "Sesi Microsoft Graph Aktif." -ForegroundColor Green

## -----------------------------------------------------------------------
## 3. PEMILIHAN OPERASI DAN LISENSI 
## -----------------------------------------------------------------------
Write-Host "`n--- 2. Pemilihan Operasi ---" -ForegroundColor Blue
Write-Host "1. Assign License"
Write-Host "2. Remove License"
$operationChoice = Read-Host "Pilih nomor menu"

switch ($operationChoice) {
    "1" { $operationType = "ASSIGN" }
    "2" { $operationType = "REMOVE" }
    default { Write-Host "Pilihan tidak valid." -ForegroundColor Red; return }
}

try {
    $availableLicenses = Get-MgSubscribedSku | Select-Object SkuPartNumber, SkuId -ErrorAction Stop
    Write-Host "`nLisensi yang Tersedia:" -ForegroundColor Yellow
    [int]$index = 1
    $promptOptions = @{}
    foreach ($lic in $availableLicenses) {
        Write-Host "${index}. $($lic.SkuPartNumber)" -ForegroundColor Magenta
        $promptOptions.Add($index, $lic)
        $index++
    }
    
    $choiceInput = Read-Host "`nMasukkan nomor lisensi"
    if (-not $promptOptions.ContainsKey([int]$choiceInput)) { throw "Nomor tidak valid." }
    
    $selectedLicense = $promptOptions[[int]$choiceInput]
    $skuPartNumberTarget = $selectedLicense.SkuPartNumber
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    return
}

## -----------------------------------------------------------------------
## 5. LOGIKA UTAMA 
## -----------------------------------------------------------------------
$allResults = @()
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$userCount = 0 

Write-Host "`n--- Memulai Proses Eksekusi ---" -ForegroundColor Magenta

foreach ($entry in $users) {
    $userCount++
    $userUpn = if ($entry.UserPrincipalName) { $entry.UserPrincipalName.Trim() } else { $null }
    if ([string]::IsNullOrWhiteSpace($userUpn)) { continue }

    Write-Host "-> [${userCount}/${totalUsers}] Memproses: ${userUpn} . . ." -ForegroundColor White

    try {
        $user = Get-MgUser -UserId $userUpn -Property 'Id', 'DisplayName', 'UsageLocation' -ErrorAction Stop
        
        if ($operationType -eq "ASSIGN" -and -not $user.UsageLocation) {
            $null = Update-MgUser -UserId $user.Id -UsageLocation $defaultUsageLocation -ErrorAction Stop
        }

        $userLicense = Get-MgUserLicenseDetail -UserId $user.Id | Where-Object { $_.SkuId -eq $selectedLicense.SkuId }

        if ($operationType -eq "ASSIGN") {
            if ($userLicense) {
                $status = "ALREADY_ASSIGNED"; $reason = "Sudah memiliki lisensi."
            } else {
                $null = Set-MgUserLicense -UserId $user.Id -AddLicenses @(@{ SkuId = $selectedLicense.SkuId }) -RemoveLicenses @() -ErrorAction Stop
                $status = "SUCCESS"; $reason = "Lisensi berhasil diberikan."
            }
        } else {
            if (-not $userLicense) {
                $status = "ALREADY_REMOVED"; $reason = "User tidak memiliki lisensi ini."
            } else {
                $null = Set-MgUserLicense -UserId $user.Id -RemoveLicenses @($selectedLicense.SkuId) -AddLicenses @() -ErrorAction Stop
                $status = "SUCCESS_REMOVED"; $reason = "Lisensi berhasil dihapus."
            }
        }

        $allResults += [PSCustomObject]@{
            UserPrincipalName = $userUpn; DisplayName = $user.DisplayName; Status = $status; Reason = $reason
        }
    }
    catch {
        Write-Host "   Gagal: $($_.Exception.Message)" -ForegroundColor Red
        $allResults += [PSCustomObject]@{
            UserPrincipalName = $userUpn; DisplayName = "Error"; Status = "FAIL"; Reason = $_.Exception.Message
        }
    }
}

## -----------------------------------------------------------------------
## 6. EKSPOR HASIL
## -----------------------------------------------------------------------
if ($allResults.Count -gt 0) {
    $outputFileName = "${operationType}_License_Results_${timestamp}.csv"
    $resultsFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName
    $allResults | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    
    Write-Host "`nProses Selesai. Laporan: ${resultsFilePath}" -ForegroundColor Green
}