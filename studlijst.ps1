# creds for fhict portal
#$cred = Import-Clixml -Path C:\scripts\fontys.cred
#$vars = @{UserName=$cred.username;Password=$Cred.GetNetworkCredential().Password} #

#open Excel
$excel=New-Object -ComObject Excel.Application
$workbook=$excel.WorkBooks.Open('C:\data\s3-i-db-T-22nj.xlsx')
$ws=$workbook.WorkSheets.item(1)
$ws1=$workbook.WorkSheets(1)
$excel.Visible = $True
$MissingType = [System.Type]::Missing
$studs =@()
# kolom is column with studentname
$kolom=6
#$aantal = 10
#$begin = 2
$vakantie =@()
#user input for this semester
#$write-host "geef eerste regel student:" -NoNewline ; [int]$begin = Read-Host
$begin = 2
write-host "geef aantal studenten:" -NoNewline ; [int]$aantal = Read-Host
write-host "geef datum eerste lesdag van de eerste week: (dd-mm-yyyy): " -NoNewline ; $startdag = Read-Host
$vakantie = @()
do {
$input = (Read-Host "geef weeknummers waarin vakantie valt (weeknr's tov eerste week) - eindig met 'end'")
if ($input -ne '') {$vakantie += $input}
}
#Loop will stop when user enter 'END' as input
until ($input -eq 'end')
 
$arrayInput

$startdag2 = [DateTime]::Parse($startdag)
#$aantal = $end - $begin
$laatste = $begin + $aantal
$MissingType = [System.Type]::Missing
$WorksheetCount = $aantal

#login to fhict portal
#$login = Invoke-WebRequest https://portal.fhict.nl -SessionVariable sv # Save session in 'sv'
#$mainPage = Invoke-WebRequest https://portal.fhict.nl -WebSession $sv -Body $vars -Method Post

# add column
$ColumnSelect = $ws.Columns("A:A")
$ColumnSelect.Insert()
$ws1.Cells.item(1,1).value2="lesweken"
$ColumnSelect = $ws.Columns("L:L")
$ColumnSelect.Insert()
$ws1.columns.item(12).NumberFormat = "dd/mm/jjjj"
$ws1.Cells.item(1,12).value2="laatste contact"
 
 #make week worksheets
 $calenderweek = 1
 $sprint=1
 $aantalvakantie = 0
 $kleur=6
  $lesweek = 1
  $weekrow=2
  $laatsteweek=19
  $presentie=15
  # begin loop
  while ($lesweek -ne $laatsteweek ) {
#$begin
if ($calenderweek -eq $vakantie[$aantalvakantie]) {$startdag2 = $startdag2.AddDays(7)
                                                   $aantalvakantie = $aantalvakantie + 1
                                                   $calenderweek = $calenderweek + 1
                                                  } else
                                                 {
if ($sprint -gt 3) {$kleur = $kleur +1
                   $sprint= 1}
$datum = $startdag2.ToString("dddd dd-MM-yy")

$ws1.Cells.Item(1,$presentie).value2=$datum
$ws1.columns.item($presentie).columnWidth = 13
$weeknaam = "week"+$lesweek
$ws1.Cells.item($weekrow,1).value2=$weeknaam
$ws1.Cells.item($weekrow,1).interior.colorindex=$kleur
$ws1.Hyperlinks.Add($ws1.Cells.Item($weekrow,1),"", "'$weeknaam'!A1") | Out-Null
if (($weeknaam -notin $($workbook.worksheets).Name)){
    #Find the current last sheet
    $LastSheet = $workbook.Worksheets|Select -Last 1
    #Make a new sheet before the current last sheet so it's near the end
    $worksheet = $workbook.worksheets.add($LastSheet)
    #Name it
    $worksheet.name = "$weeknaam"
    $i = 2
   #  $Worksheet.Columns(1).Style.Numberformat = "dd/mm/yyyy"
     $worksheet.columns.item(1).NumberFormat = "dd/mm/jjjj"
   $worksheet.Cells.Item(1,1).value2="Terug naar hoofdtab"
   $worksheet.Cells.Item(1,2).value2="$weeknaam"
   $worksheet.Cells.Item(2,1).value2="Datum"
   $worksheet.Cells.Item(3,1).value2=$datum
   $worksheet.Cells.Item(2,2).value2="Notitie"
   $worksheet.Cells.Item(2,1).Font.Bold=$True
   $worksheet.Cells.Item(2,2).Font.Bold=$True
   $worksheet.Hyperlinks.Add($worksheet.Cells.Item(1,1),"", "'studenten'!A1") | Out-Null
   while ($i -lt 25 ){
   $worksheet.Cells.item($i,1).BorderAround(1,4,3)
   $worksheet.Cells.Item($i,2).BorderAround(1,4,3)
   $i = $i+1 }
   $worksheet.columns.item(1).columnWidth = 25
   $worksheet.columns.item(2).columnWidth = 150


    #Move the last sheet up one spot, making the new sheet the new effective last sheet
    $LastSheet.Move($worksheet)
        
    }
    #on to the next
    $lesweek = $lesweek + 1 
    $calenderweek = $calenderweek + 1 
    $weekrow = $weekrow + 1
    $presentie=$presentie+1
    $sprint = $sprint + 1
    $startdag2 = $startdag2.AddDays(7)
 
 }
 }
 
#Add student worksheets
$vandaag = get-date -Format m
while ($begin -ne $laatste ) {
$studs = $ws.cells.item($begin,$kolom)
$naam = $studs.Value2
$data1=$ws.cells.item($begin,9)
$data2=$ws.cells.item($begin,12)
$vooropl = $data1.value2
$sc=$data2.value2
 $ws1.name = "studenten"
 $stunr=$ws.cells.item($begin,4).value2
 #$url="https://studentenvolg.fhict.nl/fotos/$stunr.png"
 #Invoke-WebRequest -Uri $url -WebSession $SV -OutFile "C:\scripts\students\$stunr.jpg"
#$begin
if (($naam -notin $($workbook.worksheets).Name)){
    #Find the current last sheet
    $LastSheet = $workbook.Worksheets|Select -Last 1
    #Make a new sheet before the current last sheet so it's near the end
    $worksheet = $workbook.worksheets.add($LastSheet)
    #Name it
    $worksheet.name = "$naam"
    $i = 2
   #  $Worksheet.Columns(1).Style.Numberformat = "dd/mm/yyyy"
     $worksheet.columns.item(1).NumberFormat = "dd/mm/jjjj"
   $worksheet.Cells.Item(1,1).value2="$naam"
   $worksheet.Cells.Item(1,3).value2="$vooropl"
   $worksheet.Cells.Item(1,2).value2="$sc"
   $worksheet.Cells.Item(2,1).value2="Datum"
   $worksheet.Cells.Item(3,1).value2=$vandaag
   $worksheet.Cells.Item(2,2).value2="Notitie"
   $worksheet.Cells.Item(3,2).value2="Logboek aangemaakt"
   $worksheet.Cells.Item(2,1).Font.Bold=$True
   $worksheet.Cells.Item(2,2).Font.Bold=$True
  # $foto="c:\scripts\students\"+$stunr+".jpg"
   
   $worksheet.Hyperlinks.Add($worksheet.Cells.Item(1,1),"", "'studenten'!A1") | Out-Null
   while ($i -lt 25 ){
   $worksheet.Cells.item($i,1).BorderAround(1,4,3)
   $worksheet.Cells.Item($i,2).BorderAround(1,4,3)
   $i = $i+1 }
   $worksheet.columns.item(1).columnWidth = 25
   $worksheet.columns.item(2).columnWidth = 150
   
   #add foto to sheet
   #$worksheet.Shapes.AddPicture($foto,1,0,1000,0,60,80)


    #Move the last sheet up one spot, making the new sheet the new effective last sheet
    $LastSheet.Move($worksheet)
    # create hyperlink to new sheet
    $ws.Hyperlinks.Add($ws.Cells.Item($begin,$kolom),"", "'$naam'!A1") | Out-Null

    # fill column lastcontact 
    $lastcontact = "=Lookup(2;1/('$naam'!A:A<>"""");'$naam'!A:A)"
    $ws1.Cells.item($begin,12).formulalocal=$lastcontact
 
    #on to the next
    $begin = $begin + 1 
 } 
 }
  
 $workbook.SaveAs('c:\data\s3-i-db-T-22nj-new.xlsx')
 $excel.Quit()



