$wd = New-Object -ComObject Word.Application
$wd.Visible = $false

Get-ChildItem -Filter *.doc | Foreach-Object {
    
        $fullFilename= $_.FullName
        $doc = $wd.Documents.Open($fullFilename, $false, $true)
        $table1= $doc.Tables.Item(1)

        #$description = $table1.Cell(13,2).Range.Text
        $description = Get-CellValue $table1 13 2
        
        #$value = $table1.Cell(13,6).Range.Text
        $value = Get-CellValue $table1 13 6

        #$data=$table1.Cell(10,2).Range.Text
        $data= Get-CellValue $table1 10 2

        #$number=$table1.Cell(9,2).Range.Text
        $number=Get-CellValue $table1 9 2

        $_.Name +": "+$description +"|"+$value+"|"+$data+ "|"+$number
        
}


function Get-CellValue {
 param($table, $x, $y )
    $table1.Cell($x,$y).Range.Text
}


