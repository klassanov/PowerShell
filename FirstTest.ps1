#Add-Type -AssemblyName Microsoft.Office.Interop.Excel

function Get-CellValue {
 param($table, $x, $y )
 
 $x+$y
 
}



#Get-CellValue null 3 4


$wd = New-Object -ComObject Word.Application
$wd.Visible = $false

Get-ChildItem -Filter *.doc | Foreach-Object {
    
        $fullFilename= $_.FullName
        $doc = $wd.Documents.Open($fullFilename, $false, $true)
        $table1= $doc.Tables.Item(1)
        $description = $table1.Cell(13,2).Range.Text
        $value = $table1.Cell(13,6).Range.Text
        $data=$table1.Cell(10,2).Range.Text
        $number=$table1.Cell(9,2).Range.Text

        $_.Name +": "+$description +"|"+$value+"|"+$data+ "|"+$number
        
}


