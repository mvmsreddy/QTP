
RunAction "LoginAction [Resuable_Action_Test]", oneIteration,"ashish_ait","roadrash",appLogin
print appLogin
If Not appLogin Then
Reporter.ReportEvent micFail,"Login Failed","Could not login with "
Exitrun
End If






