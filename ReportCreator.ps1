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
$wsheet.Cells.Item(1,2) = "СД Класанов и сие - Автоматичен отчет ПРИХОДИ"
$wsheet.Cells.Item(3,1) = "№"
$wsheet.Cells.Item(3,2) = "Ф-ра"
$wsheet.Cells.Item(3,3) = "Дата"
$wsheet.Cells.Item(3,4) = "Описание"
$wsheet.Cells.Item(3,5) = "Приходи (лв)"


$rowNum=4

function Get-CellValue {
 param($table, $x, $y )
    return $table1.Cell($x,$y).Range.Text
}

function Clean-NonPrintableCharacters {
  param ($excel, $str)
      return $excel.WorksheetFunction.Clean($str)
}

Get-ChildItem -Filter *.doc | Foreach-Object {

        #Read
        $fullFilename= $_.FullName
        $doc = $word.Documents.Open($fullFilename, $false, $true)
        $table1= $doc.Tables.Item(1)

        $valueRowNumber=$table1.Rows.Count-5
       
        $description = Get-CellValue $table1 13 2
        $value = Get-CellValue $table1 $valueRowNumber 3
        $data= Get-CellValue $table1 10 2

        $parsedNumber=Get-CellValue $table1 9 2
        $parsedNumber = $parsedNumber -replace "№", ""
        $number = $parsedNumber.Trim()

        $_.Name +": "+$description +"|"+$value+"|"+$data+ "|"+$number

        #Write
        $colNum=1   
          
        $wsheet.Cells.Item($rowNum, $colNum) = Clean-NonPrintableCharacters $excel $number
        #$wsheet.Cells.Item($rowNum, $colNum).NumberFormat = "@"


        $wsheet.Cells.Item($rowNum, $colNum+1) = "СД Класанов и сие"
        
        $wsheet.Cells.Item($rowNum, $colNum+2) = Clean-NonPrintableCharacters $excel $data
        $wsheet.Cells.Item($rowNum, $colNum+2).NumberFormat="dd.MM.yyyy"

        $wsheet.Cells.Item($rowNum, $colNum+3) = "Превод и легализация на документи"

        $wsheet.Cells.Item($rowNum, $colNum+4) = Clean-NonPrintableCharacters $excel $value
        #$wsheet.Cells.Item($rowNum, $colNum+4).NumberFormat="0,00"
       
        $rowNum++
}

 $excel.DisplayAlerts = $false
 $ext=".xls"
 
 $usedRange = $wsheet.UsedRange	
 $usedRange.EntireColumn.AutoFit() | Out-Null


 $path=$PSScriptRoot+"\"+"Avtomatichen-Otchet-Prihodi"+ $ext
 $workbook.SaveAs($path, 1) 
 $workbook.Close($false)
 $excel.Quit()
