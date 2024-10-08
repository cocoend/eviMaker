$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

#开始行
$startRow = 3

#开始列
$startColumn = "AB"

#图片路径
$image_folder = "C:\Users\U726466\Desktop\数選バリエーション\1007\1007"

#模板路径
$template = "C:\Users\U726466\Desktop\数選バリエーション\1007/フォーマット.xlsx"

#保存路径
$save_file = "C:\Users\U726466\Desktop\数選バリエーション\1007/フォーマットtest.xlsx"

#导入图片对象名
$prefix = "コメント"

#间隔行
$row_step = 43

#图片大小
$p_with = 50.81
$p_height = 27.52
$cm2Point = 28.3465

# open excel
$workbook = $excel.Workbooks.Open($template)
$worksheet = $workbook.Sheets.Item(1)


# 継続　普通　夜間
$images = Get-ChildItem -Path $image_folder -File -Filter "$prefix*" | Sort-Object LastWriteTime

foreach($file in $images) {

    $cell = $worksheet.Range("$startColumn" + "$startRow")

    $picture = $worksheet.Pictures().Insert($file.FullName)
    
    $picture.Width = $p_with * $cm2Point
    $picture.Height = $p_height * $cm2Point

    $picture.Top = $cell.Top
    $picture.Left = $cell.Left

    $startRow = $startRow + $row_step
}


$workbook.SaveAs($save_file)
$workbook.Close()
$Excel.Quit()


[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
