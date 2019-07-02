 Set QTP = CreateObject("QuickTest.Application")

QTP.Launch
QTP.Visible = TRUE

 
 
QTP.Open "C:\Program Files (x86)\Jenkins\workspace\SpireOne\spireone\Scripts\DevOps", TRUE
  
 
 
Set qtpResultsOpt = CreateObject("QuickTest.RunResultsOptions")

qtpResultsOpt.ResultsLocation = "C:\Program Files (x86)\Jenkins\workspace\Results"
  
  QTP.Test.Run qtpResultsOpt

 
 
QTP.Test.Close

QTP.Quit
