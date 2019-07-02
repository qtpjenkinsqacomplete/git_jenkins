 Set QTP = CreateObject("QuickTest.Application")

QTP.Launch
QTP.Visible = TRUE

 
 
QTP.Open "C:\Users\E004535\Documents\SpireOne\git_jenkins\spireone\Scripts\DevOps", TRUE
  
 
 
Set qtpResultsOpt = CreateObject("QuickTest.RunResultsOptions")

qtpResultsOpt.ResultsLocation = "C:\Program Files (x86)\Jenkins\workspace\Results"
  
 
  msgbox "welcome jenkins build"

QTP.Test.Run qtpResultsOpt

 
 
QTP.Test.Close

QTP.Quit
