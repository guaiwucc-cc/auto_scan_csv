# -*- coding: utf-8 -*-
# ��ȡ�ⲿ�����ļ�
$config = Get-Content .\config.txt
$folder = $config.Trim()

# ���庯������ CSV �ļ�ת���� Excel �ļ�������
function ConvertTo-Excel {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ -PathType 'Leaf' })]
        [string]$CsvFilePath,
        [Parameter(Mandatory = $true)]
        [string]$ExcelFilePath
    )

    # ���� CSV �ļ������ݱ���
    $csv = Import-Csv $CsvFilePath

    # �������ݱ�
    $sorted = $csv | Sort-Object -Property ���祳�`��, ���n���`��, Jancode

    # ���� Excel ����
    $excel = New-Object -ComObject Excel.Application

    # ���� Excel ����
    $excel.Visible = $false

    # ���һ���µĹ�����
    $workbook = $excel.Workbooks.Add()

    # ѡ������
    $worksheet = $workbook.Worksheets.Item(1)

    # д���ͷ
    $worksheet.Cells.Item(1,1) = "���祳�`��"
    $worksheet.Cells.Item(1,2) = "���n���`��"
    $worksheet.Cells.Item(1,3) = "Jancode"
    $worksheet.Cells.Item(1,4) = "NS؜�Ӂ���"

    # д������
    $row = 2
    foreach ($item in $sorted) {
        $worksheet.Cells.Item($row,1) = $item."���祳�`��"
        $worksheet.Cells.Item($row,2) = $item."���n���`��"
        $worksheet.Cells.Item($row,3) = $item."Jancode"
        $worksheet.Cells.Item($row,4) = $item."NS؜�Ӂ���"
        $row++
    }

    # ���� Excel �ļ�
    $workbook.SaveAs($ExcelFilePath)

    # �ͷ���Դ
    $workbook.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# ѭ������Ŀ���ļ��У������� CSV �ļ�ʱת���� Excel �ļ�
while ($true) {
    Write-Host "Scanning folder: $folder"
    Get-ChildItem $folder -Filter *.csv | ForEach-Object {
        $csvPath = $_.FullName
        $excelPath = $_.FullName.Replace(".csv", ".xlsx")
        Write-Host "Converting $csvPath to $excelPath"
        ConvertTo-Excel -CsvFilePath $csvPath -ExcelFilePath $excelPath
        Remove-Item $csvPath
    }
    Start-Sleep -Seconds 300 # �ȴ������
}