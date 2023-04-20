# �ݒ�t�@�C������t�H���_�[�p�X���擾����
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$config = Get-Content "$scriptPath\config.txt"
$folder = $config.Trim()

# log�t�@�C���̏���������
$startTime = Get-Date
$logFileName = (Get-Date).ToString('yyyy-MM-dd') + '.log'
$logFilePath = "$scriptPath\$logFileName"
# ���R�[�h����������
$csvRecode = 0
$excelRecode = 0

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

    # CSV�t�@�C���̃��R�[�h���𓝌v����
    $csvRecode = ($csv | Measure-Object).Count - 1

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
    #$worksheet.Cells.Item(1,1) = "��ЃR�[�h"
    #$worksheet.Cells.Item(1,2) = "�X�܃R�[�h"
    #$worksheet.Cells.Item(1,3) = "Jancode"
    #$worksheet.Cells.Item(1,4) = "NS�̔����i"

   # �f�[�^����������
    $row = 1
    foreach ($item in $csv) {
        $worksheet.Cells.Item($row, 1).NumberFormat = "@"
        $worksheet.Cells.Item($row, 1).Value = $item."��ЃR�[�h".ToString()
    
        $worksheet.Cells.Item($row, 2).NumberFormat = "@"
        $worksheet.Cells.Item($row, 2).Value = $item."�X�܃R�[�h".ToString()
    
        $worksheet.Cells.Item($row, 3).NumberFormat = "@"
        $worksheet.Cells.Item($row, 3).Value = $item."Jancode".ToString()

        $worksheet.Cells.Item($row,4) = $item."NS�̔����i"
        $row++
    }

    $excelRecode = $row - 1

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

# �[�����������Alog�t�@�C�����L�^���鏈��
if ($processedFiles -eq 0) {
    $endTime = Get-Date
    $logMessage = "Start Time: $startTime | End Time: $endTime | No files processed."
    Write-Host $logMessage
    Add-Content -Path $logFilePath -Value $logMessage
}

