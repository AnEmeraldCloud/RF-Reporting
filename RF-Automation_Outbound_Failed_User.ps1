#------------------------------------------------------------------------------------------
#\\\Phase 1/// - Purge any previous report files before starting the new iteration.
#7/16-Reworked plain-jane removal for date differential removal.
#Remove-Item *.xls
$PurgeDelay = (Get-Date).AddDays(-1)
Get-ChildItem "C:\Datarepo" -Recurse -Include "*.xls" | Where {$_.CreationTime -lt $PurgeDelay -and $_.Name -Match "Automation"}  | Remove-Item
#------------------------------------------------------------------------------------------
#\\\Phase 2/// - Generate Report file from RightFax 
#Setting variables for the EFR Command Line
$DateVar = Get-Date -UFormat %m/%d/%Y
$StartTime = "00:00:00 AM"
$EndTime = "11:59:59 PM"
$StartVar = $DateVar.tostring() + " " + $Starttime.tostring()
$EndVar = $Datevar.tostring() + " " + $EndTime.tostring()
#Slipstream command line for new report from RightFax Server
& "C:\Program Files (x86)\RightFax\Adminutils\EnterpriseFaxReporter1.exe" -reportName "C:\Program Files (x86)\RightFax\AdminUtils\Reports\Automation_Outbound_Failed_User.rpt" -sqlServer "Quartz\SQLEXPRESS" -sqlDatabase "RightFax2" -sqlNTAuth "True" -dateStart $StartVar -dateEnd $Datevar + " " + $EndTime -paramSearchUser "AlphaTester" -outputPath " C:\DataRepo" -outputType "XLSR" -log "Verbose"
#6/22-Waits for report to finish within 30 sec before moving forward.
#Wait-Process -Name EnterpriseFaxReporter -Timeout 30
#7/12-Changed this to the Start-Sleep command instead of the Wait-Process due to delay issues in the process spawning. 
#7/17-Changed timer to 20s from 60s as the report doesn't take long to generate < 10s whilst keeping a buffer.
Start-Sleep -s 20
#------------------------------------------------------------------------------------------
#\\\Phase 3/// - Sending report via email
$From = "Sender@Quartz"
$To = "Tester@Quartz"
$Attachment = get-childitem -Name -Filter *.xls
Send-MailMessage -From $From -To $To -Subject "Daily RightFax Report for failed faxes." -Body "Attached is the RightFax report in XLS" -Attachments $Attachment -dno onFailure -SmtpServer Quartz
#------------------------------------------------------------------------------------------