 Set QTP = CreateObject("QuickTest.Application")
QTP.Launch
QTP.Visible = TRUE
 
 
QTP.Open "C:\Users\E004535\Desktop\Spire Project\Testscripts\GUITest1", TRUE  
 
 
Set qtpResultsOpt = CreateObject("QuickTest.RunResultsOptions")
qtpResultsOpt.ResultsLocation = "C:\Users\E004535\Desktop\Spire Project\Results"  
 
 
QTP.Test.Run qtpResultsOpt
 
 
QTP.Test.Close
QTP.Quit
