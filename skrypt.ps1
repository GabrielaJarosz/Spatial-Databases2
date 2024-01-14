<#
.SYNOPSIS
Ten skrypt wykonuje serię zadań związanych z przetwarzaniem danych.

.DESCRIPTION
Skrypt pobiera plik ZIP z internetu, rozpakowuje go, przetwarza dane, zapisuje wyniki do bazy danych,
i wykonuje inne zadania związane z przetwarzaniem danych. Zdarzenia w trakcie działania skryptu są
rejestrowane w pliku log.

.CHANGELOG
2024-01-03: Utworzenie skryptu - podstawowa funkcjonalność pobierania i przetwarzania danych.
2024-01-12: Dodanie logowania zdarzeń.
2024-01-14: Poprawki błędów i optymalizacja kodu.
#>

# Funkcja do logowania
function Log-Step {
    param (
        [string]$step,
        [string]$description,
        [bool]$success
    )
    $timestamp = Get-Date -Format "yyyyMMddHHmmss"
    $status = if ($success) { "Successful" } else { "Failed" }
    $logMessage = "$timestamp – $step - $description - $status"
    Add-Content -Path $script:logFilePath -Value $logMessage
}

# Ustawienie parametryzowanych wartości
$indexNumber = "YOUR_INDEX_NUMBER"
$todayDate = Get-Date -Format "MMddyyyy"
$logName = "nazwa_logu"
$fileUrl = "ADRES_PLIKU"
$sqlHostname = "SQL_HOSTNAME"
$sqlUserId = "SQL_USERID"

# Definiowanie ścieżki pliku logowania
$timestampForLogFile = Get-Date -Format "MMddyyyy"
$processedDirectory = "D:\studia\sem2_mgr\BDP2\skrypt\PROCESSED"
$script:logFilePath = "$processedDirectory\InternetSales_new${timestampForLogFile}.log"

# Tworzenie katalogu PROCESSED jeśli nie istnieje
if (-not (Test-Path $processedDirectory)) {
    New-Item -ItemType Directory -Path $processedDirectory
}

# Prośba o wpisanie hasła
$archivePassword = Read-Host -AsSecureString "Wpisz hasło do archiwum"
$sqlPassword = Read-Host -AsSecureString "Wpisz hasło SQL"

# a,b: Pobieranie i rozpakowywanie pliku ZIP
$zipFileUrl = "http://home.agh.edu.pl/~wsarlej/dyd/bdp2/materialy/cw10/InternetSales_new.zip"
$destinationPath = "D:\studia\sem2_mgr\BDP2\skrypt"
$zipPassword = "bdp2agh"
$7zipExe = "D:\pobrane\7-Zip\7z.exe"

try {
    Invoke-WebRequest -Uri $zipFileUrl -OutFile "$destinationPath\InternetSales_new.zip"
    Start-Process -FilePath $7zipExe -ArgumentList "e", "-o$destinationPath", "-p$zipPassword", "$destinationPath\InternetSales_new.zip" -Wait
    Log-Step -step "a,b" -description "Pobieranie i rozpakowywanie pliku ZIP" -success $true
} catch {
    Log-Step -step "a,b" -description "Pobieranie i rozpakowywanie pliku ZIP" -success $false
    Write-Error "Błąd podczas pobierania i rozpakowywania pliku ZIP"
    exit
}

# c: Przetwarzanie danych z pliku
$txtPath = "$destinationPath\InternetSales_new.txt"
$csvPath = "$destinationPath\InternetSales_new.csv"
$outFileBad = "$destinationPath\InternetSales_new.bad_$(Get-Date -Format 'MMddyyyy').csv"

try {
    Get-Content $txtPath | Out-File $csvPath -Encoding UTF8
    $csvData = Import-Csv -Path $csvPath -Delimiter '|'
    $correctRows = @()
    $badRows = @()

     foreach ($row in $csvData) {
        $rowValues = $row.PSObject.Properties.Value
        if ($rowValues.Count -eq 7 -and -not [string]::IsNullOrWhiteSpace($rowValues -join '')) {
            $isRowValid = $true

            # Walidacja kolumny OrderQuantity
            if (-not ([int]::TryParse($row.OrderQuantity, [ref]0) -and $row.OrderQuantity -le 100 -and $row.OrderQuantity -gt 0)) {
                $isRowValid = $false
            }

            # Walidacja Customer_Name
            $customerNameParts = $row.Customer_Name -replace '"', '' -split ',', 2
            if ($customerNameParts.Count -ne 2 -or [string]::IsNullOrWhiteSpace($customerNameParts[0]) -or [string]::IsNullOrWhiteSpace($customerNameParts[1])) {
                $isRowValid = $false
            } else {
                $firstName = $customerNameParts[1].Trim()
                $lastName = $customerNameParts[0].Trim()
            }

            # Usunięcie wartości SecretCode
            $row.SecretCode = $null

            if ($isRowValid) {
                # Utworzenie obiektu z nowymi kolumnami
                $correctRow = [PSCustomObject]@{
                    ProductKey = $row.ProductKey
                    CurrencyAlternateKey = $row.CurrencyAlternateKey
                    LastName = $lastName
                    FirstName = $firstName
                    OrderDateKey = $row.OrderDateKey
                    OrderQuantity = $row.OrderQuantity
                    UnitPrice = $row.UnitPrice
                }

                # Dodanie wiersza do listy poprawnych
                $correctRows += $correctRow
            } else {
                # Dodanie wiersza do listy błędnych
                $badRows += $row
            }
        } else {
            # Dodanie wiersza do listy błędnych
            $badRows += $row
        }
    }

    $correctRows | Export-Csv -Path $csvPath -Delimiter '|' -NoTypeInformation -Encoding UTF8
    $badRows | ForEach-Object { $_.PSObject.Properties.Value -join '|' } | Out-File -FilePath $outFileBad -Encoding UTF8
    Log-Step -step "c" -description "Przetwarzanie danych z pliku" -success $true
} catch {
    Log-Step -step "c" -description "Przetwarzanie danych z pliku" -success $false
    Write-Error "Błąd podczas przetwarzania danych z pliku"
    exit
}

# d: Utworzenie tabeli w bazie danych
$databaseHost = "localhost"
$databaseName = "postgres"
$databaseUser = "postgres"
$psqlPath = "C:\Users\Gaba\psql.exe"
$indexNumber = "403548"
$newTable = "CUSTOMERS_$indexNumber"
$mlz = 255

try {
    $credential = Get-Credential -Message "Podaj hasło do bazy danych"
    $createTableSql = "CREATE TABLE $newTable (ProductKey INT, CurrencyAlternateKey VARCHAR($mlz), LastName VARCHAR($mlz), FirstName VARCHAR($mlz), OrderDateKey VARCHAR($mlz), OrderQuantity INT, UnitPrice VARCHAR($mlz), SecretCode VARCHAR($mlz));"
    & $psqlPath -h $databaseHost -d $databaseName -U $databaseUser -c $createTableSql -W
    Log-Step -step "d" -description "Tworzenie tabeli w bazie danych" -success $true
} catch {
    Log-Step -step "d" -description "Tworzenie tabeli w bazie danych" -success $false
    Write-Error "Błąd podczas tworzenia tabeli w bazie danych"
    exit
}

# e: Wczytywanie danych do bazy danych
try {
     # Ścieżka do przetworzonego pliku CSV
     $processedCsvPath = "$destinationPath\InternetSales_new.csv"

     # Czytanie danych z pliku CSV
     $csvData = Import-Csv -Path $processedCsvPath -Delimiter '|'
     $insertQueries = @()
 
     foreach ($row in $csvData) {
         $productKey = $row.ProductKey
         $currencyAlternateKey = $row.CurrencyAlternateKey -replace "'", "''"
         $lastName = $row.LastName -replace "'", "''"
         $firstName = $row.FirstName -replace "'", "''"
         $orderDateKey = $row.OrderDateKey -replace "'", "''"
         $orderQuantity = $row.OrderQuantity
         $unitPrice = $row.UnitPrice -replace ",", "."
         $secretCode = $row.SecretCode -replace "'", "''"
 
         # Tworzenie zapytania INSERT dla każdego wiersza
         $insertQuery = "('$productKey', '$currencyAlternateKey', '$lastName', '$firstName', '$orderDateKey', $orderQuantity, '$unitPrice', '$secretCode')"
         $insertQueries += $insertQuery
     }
 
     # Połączenie wszystkich zapytań INSERT w jedno
     $allInsertQueries = $insertQueries -join ", "
     $fullInsertSql = "INSERT INTO $newTable (ProductKey, CurrencyAlternateKey, LastName, FirstName, OrderDateKey, OrderQuantity, UnitPrice, SecretCode) VALUES $allInsertQueries;"
 
     # Zapisanie zapytania do pliku tymczasowego
     $tempSqlFile = "$env:TEMP\tempQuery.sql"
     $fullInsertSql | Out-File -FilePath $tempSqlFile
 
     # Uruchomienie psql z plikiem SQL jako wejściem
     & $psqlPath -h $databaseHost -d $databaseName -U $databaseUser -f $tempSqlFile -W
 
     # Usunięcie pliku tymczasowego
     Remove-Item $tempSqlFile
 
     Write-Host "Wczytanie danych z pliku do bazy zakończone."
     
    Log-Step -step "e" -description "Wczytywanie danych do bazy danych" -success $true
} catch {
    Log-Step -step "e" -description "Wczytywanie danych do bazy danych" -success $false
    Write-Error "Błąd podczas wczytywania danych do bazy"
    exit
}

# f: Przeniesienie przetworzonego pliku
try {
    if (-not (Test-Path $processedDirectory)) {
        New-Item -ItemType Directory -Path $processedDirectory
    }
    $processedFilePath = "$processedDirectory\${timestamp}_InternetSales_new.csv"
    Move-Item -Path $csvPath -Destination $processedFilePath
    Log-Step -step "f" -description "Przeniesienie przetworzonego pliku" -success $true
} catch {
    Log-Step -step "f" -description "Przeniesienie przetworzonego pliku" -success $false
    Write-Error "Błąd podczas przenoszenia przetworzonego pliku"
    exit
}

# g: Aktualizacja kolumny SecretCode
try {
    $randomString = Generate-RandomString
    $updateSql = "UPDATE $newTable SET SecretCode = '$randomString';"
    & $psqlPath -h $databaseHost -d $databaseName -U $databaseUser -c $updateSql -W
    Log-Step -step "g" -description "Aktualizacja kolumny SecretCode" -success $true
} catch {
    Log-Step -step "g" -description "Aktualizacja kolumny SecretCode" -success $false
    Write-Error "Błąd podczas aktualizacji kolumny SecretCode"
    exit
}

# h: Eksport zawartości tabeli do pliku CSV
try {
    $exportedCsvPath = "$processedDirectory\${timestamp}_ExportedData.csv"
    $exportSql = "\copy (SELECT * FROM $newTable) TO '$exportedCsvPath' WITH CSV HEADER;"
    & $psqlPath -h $databaseHost -d $databaseName -U $databaseUser -c $exportSql -W
    Log-Step -step "h" -description "Eksport zawartości tabeli do pliku CSV" -success $true
} catch {
    Log-Step -step "h" -description "Eksport zawartości tabeli do pliku CSV" -success $false
    Write-Error "Błąd podczas eksportowania zawartości tabeli"
    exit
}

# i: Kompresowanie pliku CSV
try {
    $zipPath = "$processedDirectory\${timestamp}_ExportedData.zip"
    Compress-Archive -Path $exportedCsvPath -DestinationPath $zipPath
    Log-Step -step "i" -description "Kompresowanie pliku CSV" -success $true
} catch {
    Log-Step -step "i" -description "Kompresowanie pliku CSV" -success $false
    Write-Error "Błąd podczas kompresowania pliku CSV"
    exit
}
