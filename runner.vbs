 Set QTP = CreateObject("QuickTest.Application")
QTP.Launch
QTP.Visible = TRUE
 
 
QTP.Open "C:\Users\E004535\Desktop\Spire Project\Testscripts\GUITest1", TRUE  
 
 
Set qtpResultsOpt = CreateObject("QuickTest.RunResultsOptions")
qtpResultsOpt.ResultsLocation = "C:\Program Files (x86)\Jenkins\workspace\Results"  
 
  msgbox "welcome jenkins build"
QTP.Test.Run qtpResultsOpt
 
 
QTP.Test.Close
QTP.Quit
