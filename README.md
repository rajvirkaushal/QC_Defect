# QC_Defect
Dim qcURL 
Dim qcID 
Dim qcPWD 
Dim qcDomain 
Dim qcProject 
Dim tdConnection 
Dim BugFactory 
Dim Bug 


   qcURL = "http://xxxxxxxxx:8080/qcbin"
   qcID = "xxxxxxxx"
   qcPWD = "xxxxxxx"
   qcDomain = "xxxxxxxx"
   qcProject = "xxxxxxxx"


On Error Resume NEXT

' read XML result File 

Dim xmlDoc, objNodeList, plot
Dim arrDefects(10,2)


Set xmlDoc = CreateObject("Msxml2.DOMDocument")

xmlDoc.setProperty "SelectionLanguage", "XPath"

xmlDoc.load("C:\Jenkins\workspace\2. Smoke Test_TOSCA\result.xml")

Set suitenode = xmlDoc.getElementsByTagName("testsuite")

Set objNodeList = xmlDoc.getElementsByTagName("testcase")

Set failureList = xmlDoc.getElementsByTagName("failure")
plot="No Value"

for each n in suitenode
	buildno = n.getattribute("name")
	'wscript.echo "Build number: " & buildno
Next

FailureCount = 0 

If objNodeList.length > 0 then
    	          For each x in objNodeList
				     JobName=x.getattribute("name")	
					 timestamp=x.getAttribute("timestamp")
					
					 
					 for each y in x.ChildNodes
					 failurename = y.getAttribute("message")
					 'msgbox  "Test case name: " & JobName & "  " & "  Failure message: " & y.text & "  Timestamp: " & timestamp
				     	 arrDefects(FailureCount,0) = JobName 
					 arrDefects(FailureCount,1) = y.text
					 arrDefects(FailureCount,2) = buildno

					 'msgbox arrDefects(FailureCount,0) & arrDefects(FailureCount,1) & arrDefects(FailureCount,2)
					 FailureCount = FailureCount + 1


					 Next
Next
Else
    'msgbox chr(34) & "failure" & chr(34) & " field not found."
End If





'Display a message in Status bar
 'Msgbox  "Connecting to Quality Center.. Wait..."

' Create a Connection object to connect to Quality Center
 Set tdConnection = CreateObject("TDApiOle80.TDConnection")
	'Msgbox Err.description 

tdConnection.InitConnectionEx qcURL

'Authenticating with username and password
   tdConnection.Login qcID, qcPWD
		
'connecting to the domain and project
   tdConnection.Connect qcDomain, qcProject
 
If (TDConnection.LoggedIn <> True) Then
	wscript.echo "QC User Authentication Failed"
Else 
	wscript.echo "Authentication Successful"

	'Get the IBugFactory 
	'Set BugFactory = CreateObject("TDApiOle80.TDConnection")
	Set BugFactory = tdConnection.BugFactory  
	'Msgbox Err.description
	'Msgbox "Bug factory object created"
	
	'Rasing defects in QC 

	iCount = 0

 For iCount = 0 to FailureCount-1

	Set Bug = BugFactory.AddItem(NULL) 
	'Msgbox Err.description
	'Msgbox "Bug object cretaed"

	'Enter values for required fields for the defect found 
	Bug.Status = "New" 
	'Msgbox "Status" & Err.description 

	Bug.Summary = arrDefects(iCount,0) 
	'Msgbox "Bug.Summary" & Err.description 

	Bug.Field("BG_DESCRIPTION") = arrDefects(iCount,1) 

	Bug.DetectedBy = "demouser02" 
	'Msgbox "Bug.DetectedBy" & Err.description

	Bug.Field("BG_SEVERITY") = "3-High"
	'Msgbox "Bug.Severity" & Err.description

	bug.Field("BG_DETECTION_DATE") = "2016-06-24"
	'Msgbox "BG_DETECTION_DATE" & Err.description

	'Bug.Field("BG_BUILD_DETECTED") = "123456"

	
	Bug.Post()

	'Msgbox Err.description 
	'Msgbox "Defect cretaed"

	wscript.echo "Defect id for  " & Bug.Summary & "  is  " & bug.Field("BG_BUG_ID")

	set attachFact = bug.Attachments
	set attachObj = attachFact.AddItem(NULL)

	attachObj.FileName = "c:\\QC\\DefectAttachment.txt"
	attachObj.Type = 1
	attachObj.Description = "Test Description"
	attachObj.Post()

	' Linking defect to test run instance 

tsFolderPath = "Root\TOSCA_SMOKE_Test"
TestSetName =  buildno
TestScriptName  = "[1]" & buildno & "_" & arrDefects(iCount,0)

Set labTreeMgr= tdConnection.testsettreemanager


Dim labFolder
Set labFolder = labTreeMgr.NodeByPath(tsFolderPath)
If labFolder Is Nothing Then
 wscript.echo "No nodes found"
End If

Set tsList = labFolder.FindTestSets(TestSetName)
If tsList.Count > 1 Then
 wscript.echo "Found more than one test set."
ElseIf tsList.Count < 1 Then
 wscript.echo "Test set not found. "
End If

'msgbox "count:-  " & tsList.Count
               
Dim theTestSet, tsFolder, tsTestList, tsTest, tsTestFactory
Dim BugLinkF, b2Test

Set theTestSet = tsList.Item(1)
                   
Set tsFolder = theTestSet.TestSetFolder
Set tsTestFactory = theTestSet.tsTestFactory
Set tsTestList = tsTestFactory.NewList("")
 
               
For Each tsTest In tsTestList
'msgbox "test case name " & tsTest.Name
 If tsTest.Name = TestScriptName Then
	wscript.echo tsTest.Name & "  test Found"
	
	Set BugLinkF = tsTest.BugLinkFactory
	Set b2Test = BugLinkF.AddItem(bug.ID)
	b2Test.LinkType = "Related"
	b2Test.Post
	if err.description <> "" then
		wscript.echo "Not able to link to test case" 
	Else 
		wscript.echo "Test case linked"
	End if 

        Exit For
 End If
Next

 Next 

End If
