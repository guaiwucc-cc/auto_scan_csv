# 設定ファイルからフォルダーパスを取得する
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$config = Get-Content "$scriptPath\config.txt"
$folder = $config.Trim()

# 実行開始時間を取得する
$startTime = Get-Date

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
    $worksheet.Cells.Item(1,1) = "会社コード"
    $worksheet.Cells.Item(1,2) = "店舗コード"
    $worksheet.Cells.Item(1,3) = "Jancode"
    $worksheet.Cells.Item(1,4) = "NS販売価格"

   # データを書き込む
    $row = 2
    foreach ($item in $csv) {
        $worksheet.Cells.Item($row,1) = $item."会社コード"
        $worksheet.Cells.Item($row,2) = $item."店舗コード"
        $worksheet.Cells.Item($row,3) = $item."Jancode"
        $worksheet.Cells.Item($row,4) = $item."NS販売価格"
        $row++
    }

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
Get-ChildItem $folder -Filter *.csv | ForEach-Object {
    $csvPath = $_.FullName
    $excelPath = $_.FullName.Replace(".csv", ".xlsx")
    Write-Host "Converting $csvPath to $excelPath"
    try {
        ConvertTo-Excel -CsvFilePath $csvPath -ExcelFilePath $excelPath
        Remove-Item $csvPath
        $processedFiles++
    } catch {
        Write-Host "Error converting"
    }

    $endTime = Get-Date
    $elapsedTime = $endTime - $startTime
    $logMessage = "Execution Time: $($elapsedTime.ToString()) | Converted From: $csvPath | Converted To: $excelPath | Successfully Processed Files: $processedFiles"
    Write-Host $logMessage
}

