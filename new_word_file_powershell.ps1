#Create Word File on Desktop
$date = Get-DAte -Format FileDateTime
New-Item "C:\temp\test-$date.txt"
<# $word = New-Object -ComObject Word.Application
$doc = $word.Documents.Add()
$doc.SaveAs("$env:userprofile\Desktop\testOBTScript.docx")
$word.Quit() #>
