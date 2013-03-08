 @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 4").WebCheckBox("location checkboxes[]")_;_script infofile_;_ZIP::ssf27.xml_;_
 @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 3").WebList("ip session[affiliation]")_;_script infofile_;_ZIP::ssf8.xml_;_
 @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf36.xml_;_
 @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 3").WebElement("(704.0) Alopecia, Baldness,")_;_script infofile_;_ZIP::ssf35.xml_;_

Dim Phase(2)
phase(0) = "II"
phase(1) = "III"
phase(2) = "IV"


DataTable.AddSheet("Phase II")
DataTable.AddSheet("Phase III")
DataTable.AddSheet("Phase IV")

DataTable.ImportSheet "C:\Program Files\HP\QuickTest Professional\Tests\Phases.xls", "Phase II",  "Phase II"
DataTable.ImportSheet "C:\Program Files\HP\QuickTest Professional\Tests\Phases.xls", "Phase III",  "Phase III"
DataTable.ImportSheet "C:\Program Files\HP\QuickTest Professional\Tests\Phases.xls", "Phase IV", "Phase IV"


For i=0 to 2
'DataTable.ImportSheet "C:\Program Files\HP\QuickTest Professional\Tests\Phases.xls", "Phase "&i,  "Phase "&i.name
pht="Phase "&phase(i)
cntr=DataTable.GetSheet("Phase "&phase(i)).GetRowCount

'cntr="Phase "&phase(i).GetRowCount 
For j=1to cntr @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager").Link("My Forecasts")_;_script infofile_;_ZIP::ssf1.xml_;_
DataTable.GetSheet("Phase "&phase(i)).SetCurrentRow(j) 


Browser("Grants Manager").Page("Grants Manager").Link("My Forecasts").Click @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 2").Link("Create New Forecast")_;_script infofile_;_ZIP::ssf2.xml_;_
Browser("Grants Manager").Page("Grants Manager_2").Link("Create New Forecast").Click @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 3").WebList("ip session[phase id]")_;_script infofile_;_ZIP::ssf3.xml_;_
Browser("Grants Manager").Page("Grants Manager_3").Sync
wait 3
Browser("Grants Manager").Page("Grants Manager_3").WebList("ip_session_phase_id").Select "Phase "&phase(i) @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 3").Link("Click to change the Indication")_;_script infofile_;_ZIP::ssf4.xml_;_
'blnk=Browser("Grants Manager").Page("Grants Manager_3").WebList("ip_session_phase_id").GetTOProperty("selected item index")
'If blnk  <2 Then
'     Do Until bnk =>2
'	Browser("Grants Manager").Page("Grants Manager_3").WebList("ip_session_phase_id").Select "Phase "&phase(i) @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 3").Link("Click to change the Indication")_;_script infofile_;_ZIP::ssf4.xml_;_
'	wait 1
'	blnk=Browser("Grants Manager").Page("Grants Manager_3").WebList("ip_session_phase_id").GetROProperty("selectedIndex")
'	 Loop
'End If @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 3").Link("Click to change the Indication")_;_script infofile_;_ZIP::ssf4.xml_;_
wait 3

Browser("Grants Manager").Page("Grants Manager_3").Sync
Browser("Grants Manager").Page("Grants Manager_3").Link("Click to change the Indication").Click @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 3").WebEdit("search indication term")_;_script infofile_;_ZIP::ssf5.xml_;_
Browser("Grants Manager").Page("Grants Manager_3").WebEdit("search_indication_term").Set DataTable("Indication_Code", "Phase "&phase(i)) @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 3").WebButton("Search")_;_script infofile_;_ZIP::ssf6.xml_;_

Browser("Grants Manager").Page("Grants Manager_3").WebButton("Search").Click @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 3").WebElement("(009.2) Infectious Diarrhea")_;_script infofile_;_ZIP::ssf7.xml_;_
Browser("Grants Manager").Page("Grants Manager_3").Sync
wait 2
Browser("Grants Manager").Page("Grants Manager_3").VirtualButton("button_2").Click @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf41.xml_;_

Browser("Grants Manager").Page("Grants Manager_3").WebList("ip_session[affiliation]").Select  DataTable("Site_Type", "Phase "&phase(i)) @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 3").WebList("ip session[inpatient status id")_;_script infofile_;_ZIP::ssf9.xml_;_
Browser("Grants Manager").Page("Grants Manager_3").WebList("ip_session[inpatient_status_id").Select  DataTable("Subject_Status", "Phase "&phase(i)) @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 3").WebList("ip session[study duration id]")_;_script infofile_;_ZIP::ssf10.xml_;_
Browser("Grants Manager").Page("Grants Manager_3").WebList("ip_session[study_duration_id]").Select  DataTable("Duration", "Phase "&phase(i)) @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 3").Link("Next")_;_script infofile_;_ZIP::ssf11.xml_;_
Browser("Grants Manager").Page("Grants Manager_3").Link("Next").Click @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 4").WebCheckBox("location checkboxes[]")_;_script infofile_;_ZIP::ssf12.xml_;_

Browser("Grants Manager").Page("Grants Manager_4").WebElement(DataTable("Country", "Phase "&phase(i))).click @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 4").WebCheckBox("location checkboxes[]")_;_script infofile_;_ZIP::ssf27.xml_;_
Browser("Grants Manager").Page("Grants Manager_4").Link("Add Selected Countries").Click @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 4").Link("Next")_;_script infofile_;_ZIP::ssf16.xml_;_
Browser("Grants Manager").Page("Grants Manager_4").Sync
wait 2
Browser("Grants Manager").Page("Grants Manager_4").Link("Next").Click

val_lcps=""
val_mcps=""
val_hcps=""
val_lcpv=""
val_mcpv=""
val_hcpv=""

val_lcps=Browser("Grants Manager").Page("Grants Manager_5").WebElement("Low_cps").GetROProperty("innerText")
DataTable.Value("Low_cps","Phase "&phase(i)) =val_lcps

val_mcps=Browser("Grants Manager").Page("Grants Manager_5").WebElement("Med_cps").GetROProperty("innerText")
DataTable.Value("Middle_cps","Phase "&phase(i)) =val_mcps


val_hcps=Browser("Grants Manager").Page("Grants Manager_5").WebElement("High_cps").GetROProperty("innerText")
DataTable.Value("High_cps","Phase "&phase(i)) = val_hcps


val_lcpv=Browser("Grants Manager").Page("Grants Manager_5").WebElement("Low_cpv").GetROProperty("innerText")
DataTable.Value("Low_cpv","Phase "&phase(i)) =val_lcpv

val_mcpv=Browser("Grants Manager").Page("Grants Manager_5").WebElement("Med_cpv").GetROProperty("innerText")
DataTable.Value("Middle_cpv","Phase "&phase(i)) =val_mcpv

val_hcpv=Browser("Grants Manager").Page("Grants Manager_5").WebElement("High_cpv").GetROProperty("innerText")
DataTable.Value("High_cpv","Phase "&phase(i)) =val_hcpv


Browser("Grants Manager").Page("Grants Manager_5").Sync

Browser("Grants Manager").Page("Grants Manager_5").Link("Trial List").Click @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 5").Link("Trial List")_;_script infofile_;_ZIP::ssf37.xml_;_
Browser("Grants Manager").Page("Grants Manager_5").Link("Don't Save & Leave").Click @@ hightlight id_;_Browser("Grants Manager").Page("Grants Manager 5").Link("Don't Save & Leave")_;_script infofile_;_ZIP::ssf38.xml_;_
Next



'export data sheet
DataTable.ExportSheet "C:\Program Files\HP\QuickTest Professional\Tests\phases.xls","Phase "&phase(i)




Next





































