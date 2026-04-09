$ErrorActionPreference = 'Stop'

$workspace = 'C:\Users\admin\Documents\New project'
$sourceImage = 'E:\xwechat_files\wxid_zi4b1lrkk46622_9871\temp\RWTemp\2026-04\fa52da95f686b29dbe40ca5c0e05f484.jpg'
$cropImage = Join-Path $workspace '从业人员基本情况登记表-身份证区域.jpg'
$docxPath = Join-Path $workspace '从业人员基本情况登记表.docx'
$pdfPath = Join-Path $workspace '从业人员基本情况登记表.pdf'

function Set-CellText {
    param(
        [Parameter(Mandatory = $true)] $Cell,
        [Parameter(Mandatory = $true)] [string] $Text,
        [int] $Align = 1,
        [int] $FontSize = 12,
        [string] $FontName = '宋体'
    )

    $range = $Cell.Range
    $range.End = $range.End - 1
    $range.Text = $Text
    $range.Font.Name = $FontName
    $range.Font.Size = $FontSize
    $range.ParagraphFormat.Alignment = $Align
}

function Set-TableFormatting {
    param(
        [Parameter(Mandatory = $true)] $Table
    )

    $Table.Borders.Enable = 1
    $Table.Range.Font.Name = '宋体'
    $Table.Range.Font.Size = 12
    $Table.Rows.LeftIndent = 0
    $Table.Rows.Alignment = 1
    $Table.TopPadding = 6
    $Table.BottomPadding = 6
    $Table.LeftPadding = 4
    $Table.RightPadding = 4

    $widths = @(55, 72, 60, 54, 58, 75, 104)
    for ($i = 1; $i -le $widths.Count; $i++) {
        $Table.Columns.Item($i).Width = $widths[$i - 1]
    }

    $rowHeights = @(46, 54, 48, 48, 46, 360)
    for ($i = 1; $i -le $rowHeights.Count; $i++) {
        $Table.Rows.Item($i).HeightRule = 2
        $Table.Rows.Item($i).Height = $rowHeights[$i - 1]
    }
}

Add-Type -AssemblyName System.Drawing
$bitmap = [System.Drawing.Bitmap]::FromFile($sourceImage)
try {
    $rect = New-Object System.Drawing.Rectangle(400, 760, 560, 680)
    $crop = $bitmap.Clone($rect, $bitmap.PixelFormat)
    try {
        $crop.Save($cropImage, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    }
    finally {
        $crop.Dispose()
    }
}
finally {
    $bitmap.Dispose()
}

$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Add()

    $pageSetup = $doc.PageSetup
    $pageSetup.TopMargin = 56.7
    $pageSetup.BottomMargin = 56.7
    $pageSetup.LeftMargin = 56.7
    $pageSetup.RightMargin = 56.7

    $selection = $word.Selection
    $selection.Font.Name = '宋体'
    $selection.Font.Size = 12
    $selection.ParagraphFormat.Alignment = 1
    $selection.TypeText('从业人员基本情况登记表')
    $selection.Font.Size = 18
    $selection.Range.Font.Size = 18
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $range = $doc.Bookmarks.Item('\endofdoc').Range
    $table = $doc.Tables.Add($range, 6, 7)
    Set-TableFormatting -Table $table

    $table.Cell(2, 3).Merge($table.Cell(2, 4))
    $table.Cell(2, 4).Merge($table.Cell(2, 5))
    $table.Cell(3, 2).Merge($table.Cell(3, 6))
    $table.Cell(4, 2).Merge($table.Cell(4, 6))
    $table.Cell(2, 7).Merge($table.Cell(4, 7))
    $table.Cell(5, 1).Merge($table.Cell(5, 2))
    $table.Cell(5, 2).Merge($table.Cell(5, 3))
    $table.Cell(5, 5).Merge($table.Cell(5, 6))
    $table.Cell(6, 1).Merge($table.Cell(6, 7))

    Set-CellText $table.Cell(1, 1) '姓  名'
    Set-CellText $table.Cell(1, 2) '赵苗娟'
    Set-CellText $table.Cell(1, 3) '性  别'
    Set-CellText $table.Cell(1, 4) '女'
    Set-CellText $table.Cell(1, 5) '年  龄'
    Set-CellText $table.Cell(1, 6) '31'
    Set-CellText $table.Cell(1, 7) ''

    Set-CellText $table.Cell(2, 1) '职务'
    Set-CellText $table.Cell(2, 2) '教师'
    Set-CellText $table.Cell(2, 3) '工作岗位'
    Set-CellText $table.Cell(2, 4) '东沟中心校'
    Set-CellText $table.Cell(2, 5) '照片'

    Set-CellText $table.Cell(3, 1) '籍贯'
    Set-CellText $table.Cell(3, 2) '山西省乡宁县双鹤乡社里村'

    Set-CellText $table.Cell(4, 1) '住址'
    Set-CellText $table.Cell(4, 2) '建行宿舍'

    Set-CellText $table.Cell(5, 1) '联系电话'
    Set-CellText $table.Cell(5, 2) '15343547970'
    Set-CellText $table.Cell(5, 3) '身份证号'
    Set-CellText $table.Cell(5, 5) '142631199208064523'

    $imageCellRange = $table.Cell(6, 1).Range
    $imageCellRange.End = $imageCellRange.End - 1
    $imageCellRange.ParagraphFormat.Alignment = 1
    $shape = $doc.InlineShapes.AddPicture($cropImage, $false, $true, $imageCellRange)
    $shape.LockAspectRatio = -1
    $shape.Width = 360

    $doc.SaveAs([ref] $docxPath, [ref] 16)
    $doc.ExportAsFixedFormat($pdfPath, 17)

    Write-Output "DOCX=$docxPath"
    Write-Output "PDF=$pdfPath"
    Write-Output "CROP=$cropImage"
}
finally {
    if ($doc -ne $null) {
        $doc.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
    }
    if ($word -ne $null) {
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
