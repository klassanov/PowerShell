$word = New-Object -ComObject Word.Application
$word.Visible = $false

$excel = New-Object -ComObject Excel.application
$excel.Visible = $false

#$fixedDescription="Превод и легализация на документи"
$currentYear=get-date -Format yyyy
$previousYear=$currentYear-1
$workbook = $excel.Workbooks.Add()
$wsheet= $workbook.Worksheets.Item(1)
#$wsheet.Name = "Fakturi"
$wsheet.Cells.Item(1,2) = "СД Класанов и сие - ПРИХОДИ за $previousYear година"
$wsheet.Cells.Item(3,1) = "№"
$wsheet.Cells.Item(3,2) = "Ф-ра"
$wsheet.Cells.Item(3,3) = "Дата"
$wsheet.Cells.Item(3,4) = "Описание"
$wsheet.Cells.Item(3,5) = "Приходи(лв)"

Get-ChildItem -Filter *.doc | Foreach-Object {
    
        #Read
        $fullFilename= $_.FullName
        $doc = $word.Documents.Open($fullFilename, $false, $true)
        $table1= $doc.Tables.Item(1)

       
        $description = Get-CellValue $table1 13 2
        $value = Get-CellValue $table1 13 6
        $data= Get-CellValue $table1 10 2
        $number=Get-CellValue $table1 9 2
        $_.Name +": "+$description +"|"+$value+"|"+$data+ "|"+$number

        #Write     
        $directory=$_.DirectoryName
        
}

 $excel.DisplayAlerts = 'False'
 $ext=".xls"
 $path=$PSScriptRoot+"\"+"OTCHET "+ $currentYear+ $ext
 $path
 $workbook.SaveAs($path, 1) 
 $workbook.Close
 $excel.Quit()



function Get-CellValue {
 param($table, $x, $y )
    $table1.Cell($x,$y).Range.Text
}


