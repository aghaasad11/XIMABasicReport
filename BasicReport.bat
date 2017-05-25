:loop
if exist X:\BasicReportBilling\*.pdf (
cscript C:\Scripts\XIMA\BasicReport\BasicReportBillingSendGrid.vbs
cscript C:\Scripts\XIMA\BasicReport\BasicReportL1SendGrid.vbs
cscript C:\Scripts\XIMA\BasicReport\GroupPerformanceL1.vbs
)
move /Y X:\BasicReportBilling\Archive\*.pdf X:\BasicReportBilling\Archive\Archive
move /Y "X:\Basic Report\Archive\*.pdf" "X:\Basic Report\Archive\Archive"
cscript C:\Scripts\XIMA\BasicReport\GroupPerformanceL1.vbs
timeout /T 300
Goto loop
