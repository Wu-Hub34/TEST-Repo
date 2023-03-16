#Create Word File on Desktop
New-Item "$env:userprofile\Desktop\test.txt"
<# $word = New-Object -ComObject Word.Application
$doc = $word.Documents.Add()
$doc.SaveAs("$env:userprofile\Desktop\testOBTScript.docx")
$word.Quit() #>
