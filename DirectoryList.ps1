Get-ChildItem -Filter *.doc | Foreach-Object {
    
    $_.Name

}




Read-Host -Prompt "Press Enter to exit"
