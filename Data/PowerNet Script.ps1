Set-ExecutionPolicy -Scope Process Unrestricted
#http://stackoverflow.com/questions/16460163/ps1-cannot-be-loaded-because-the-execution-of-scripts-is-disabled-on-this-syste
#Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
#Set-ExecutionPolicy Unrestricted # Needs an elevated shell
#Set-ExecutionPolicy -Scope LocalMachine Unrestricted


#$xl = new-object -comobject excel.application 
$xl = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
$xl.Visible = $true

$ChartNames = "my chart 1", "y", "my chart 2"
$ChartObjects = @()
$wb = $xl.workbooks.open("C:\Users\sdo\Documents\SAIDI and SAIFI\SAIDI SAIFI Calculator.xlsm")
Start-Sleep -m 20 # Wait for Excel to open
$ws = $wb.worksheets.item("Calculation User Defined")
$charts = $ws.ChartObjects()
For ($i=1; $i -le $charts.Count; $i++) {
     #If ($ChartNames -contains($charts.Item($i).Chart.ChartTitle.Text)) {
     #   $charts.Item($i).Copy()
     #}
     # Copy all the charts (overwrites the old array and adds new value)
     $ChartObjects += $charts.Item($i)
}

$wd = new-object -comobject Word.application
$wd.visible = $true
$doc = $wd.documents.open("C:\Users\sdo\Documents\SAIDI and SAIFI\Report Template.docx")
Start-Sleep -m 20 # Wait for Word to open

$default = [Type]::Missing
Foreach ($chart in $ChartObjects) {
    $chart.Activate()
    $chart.Copy()
    #$doc.Content.Paste() # replaces the first copy
    #$wd.Selection.Paste()
    $wd.Selection.PasteSpecial($default, $default, $default, $default, 9, $default, $default)
}