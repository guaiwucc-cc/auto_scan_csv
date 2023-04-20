# 設定ファイルからフォルダーパスを取得する
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$config = Get-Content "$scriptPath\config.txt"
$folder = $config.Trim()

# logファイルの処理初期化
$startTime = Get-Date
$logFileName = (Get-Date).ToString('yyyy-MM-dd') + '.log'
$logFilePath = "$scriptPath\$logFileName"
# レコード数を初期化
$csvRecode = 0
$excelRecode = 0

# 変換関数を定義する
function ConvertTo-Excel {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ -PathType 'Leaf' })]
        [string]$CsvFilePath,
        [Parameter(Mandatory = $true)]
        [string]$ExcelFilePath
    )

    # CSVファイルを読み込む
    $csv = Import-Csv $CsvFilePath -Header '会社コード', '店舗コード', 'Jancode', 'NS販売価格'

    # CSVファイルのレコード数を統計する
    $csvRecode = ($csv | Measure-Object).Count - 1

    # Jancodeでソートされたデータを取得する
    $sorted = $csv | Sort-Object -Property Jancode

    # Excelアプリケーションオブジェクトを作成する
    $excel = New-Object -ComObject Excel.Application

    # Excelを非表示にする
    $excel.Visible = $false

    # 新しいワークブックを作成する
    $workbook = $excel.Workbooks.Add()

    # 最初のワークシートオブジェクトを取得する
    $worksheet = $workbook.Worksheets.Item(1)

    # ヘッダーを書き込む
    #$worksheet.Cells.Item(1,1) = "会社コード"
    #$worksheet.Cells.Item(1,2) = "店舗コード"
    #$worksheet.Cells.Item(1,3) = "Jancode"
    #$worksheet.Cells.Item(1,4) = "NS販売価格"

   # データを書き込む
    $row = 1
    foreach ($item in $csv) {
        $worksheet.Cells.Item($row, 1).NumberFormat = "@"
        $worksheet.Cells.Item($row, 1).Value = $item."会社コード".ToString()
    
        $worksheet.Cells.Item($row, 2).NumberFormat = "@"
        $worksheet.Cells.Item($row, 2).Value = $item."店舗コード".ToString()
    
        $worksheet.Cells.Item($row, 3).NumberFormat = "@"
        $worksheet.Cells.Item($row, 3).Value = $item."Jancode".ToString()

        $worksheet.Cells.Item($row,4) = $item."NS販売価格"
        $row++
    }

    $excelRecode = $row - 1

    # Excelファイルを保存する
    $workbook.SaveAs($ExcelFilePath)

    # ワークブックとExcelアプリケーションを閉じる
    $workbook.Close()
    $excel.Quit()

    # Excelオブジェクトを解放する
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# フォルダーをスキャンする
Write-Host "Scanning folder: $folder"
$processedFiles = 0
$failedFiles = 0
Get-ChildItem $folder -Filter *.csv | ForEach-Object {
    $csvPath = $_.FullName
    $excelPath = $_.FullName.Replace(".csv", ".xlsx")
    Write-Host "Converting $csvPath to $excelPath"
    try {
        ConvertTo-Excel -CsvFilePath $csvPath -ExcelFilePath $excelPath
        Remove-Item $csvPath
        $processedFiles++
    } catch {
        Write-Host "Error converting $csvPath"
        $failedFiles++
    }

    $endTime = Get-Date
    $logMessage = "Start Time: $startTime | End Time: $endTime | Converted From: $csvPath | Converted To: $excelPath | csv recode: $csvRecode - excel recode: $excelRecode | Successfully Processed Files: $processedFiles | Failed Files: $failedFiles"
    Write-Host $logMessage
    Add-Content -Path $logFilePath -Value $logMessage
}

# ゼロ件処理時、logファイルを記録する処理
if ($processedFiles -eq 0) {
    $endTime = Get-Date
    $logMessage = "Start Time: $startTime | End Time: $endTime | No files processed."
    Write-Host $logMessage
    Add-Content -Path $logFilePath -Value $logMessage
}

