# �ݒ�t�@�C������t�H���_�[�p�X���擾����
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$config = Get-Content "$scriptPath\config.txt"
$folder = $config.Trim()

# log�t�@�C���̏���������
$startTime = Get-Date
$logFileName = (Get-Date).ToString('yyyy-MM-dd') + '.log'
$logFilePath = "$scriptPath\$logFileName"

# �ϊ��֐����`����
function ConvertTo-Excel {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ -PathType 'Leaf' })]
        [string]$CsvFilePath,
        [Parameter(Mandatory = $true)]
        [string]$ExcelFilePath
    )

    # CSV�t�@�C����ǂݍ���
    $csv = Import-Csv $CsvFilePath -Header '��ЃR�[�h', '�X�܃R�[�h', 'Jancode', 'NS�̔����i'

    # Jancode�Ń\�[�g���ꂽ�f�[�^���擾����
    $sorted = $csv | Sort-Object -Property Jancode

    # Excel�A�v���P�[�V�����I�u�W�F�N�g���쐬����
    $excel = New-Object -ComObject Excel.Application

    # Excel���\���ɂ���
    $excel.Visible = $false

    # �V�������[�N�u�b�N���쐬����
    $workbook = $excel.Workbooks.Add()

    # �ŏ��̃��[�N�V�[�g�I�u�W�F�N�g���擾����
    $worksheet = $workbook.Worksheets.Item(1)

    # �w�b�_�[����������
    $worksheet.Cells.Item(1,1) = "��ЃR�[�h"
    $worksheet.Cells.Item(1,2) = "�X�܃R�[�h"
    $worksheet.Cells.Item(1,3) = "Jancode"
    $worksheet.Cells.Item(1,4) = "NS�̔����i"

   # �f�[�^����������
    $row = 2
    foreach ($item in $csv) {
        $worksheet.Cells.Item($row,1) = $item."��ЃR�[�h"
        $worksheet.Cells.Item($row,2) = $item."�X�܃R�[�h"
        $worksheet.Cells.Item($row,3) = $item."Jancode"
        $worksheet.Cells.Item($row,4) = $item."NS�̔����i"
        $row++
    }

    # Excel�t�@�C����ۑ�����
    $workbook.SaveAs($ExcelFilePath)

    # ���[�N�u�b�N��Excel�A�v���P�[�V���������
    $workbook.Close()
    $excel.Quit()

    # Excel�I�u�W�F�N�g���������
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# �t�H���_�[���X�L��������
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
    Add-Content -Path $logFilePath -Value $logMessage
}

