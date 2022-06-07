'junit XML'
'https://dev.classmethod.jp/testing/unittesting/junit-xml-format/'
Class VbsUnitTest
	Private objDOM
	Private testsuiteCount
	Private testcaseCount
	Private testcaseCountGlobal
	Private failureCount
	Private failureCountGlobal
	Private errorCount
	Private errorCountGlobal
	Private testsuitesElement
	Private testsuiteElement'Current Testsuite'
	Private testcaseElement'Current Testcase'
	Private testsuitesTimer'Current Testsuites Timer'
	Private testsuiteTimer'Current Testsuite Timer
	Private testcaseTimer'Current TestCase Timer

	'Constructor
	Private Sub Class_Initialize()
        Set objDOM = CreateObject("MSXML2.DOMDocument")
		objDOM.appendChild(objDOM.createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8"""))
		Set testsuitesElement = objDOM.createElement("testsuites")
		objDOM.appendChild(testsuitesElement)
		testsuiteCount = 0
		testcaseCount = 0
		testcaseCountGlobal = 0
		failureCount = 0
		failureCountGlobal = 0
		errorCount = 0
		errorCountGlobal = 0
    End Sub
	
	'Destructor
	Private Sub Class_Terminate
		Set objDOM = Nothing
	End Sub

	'Property
	Public Property Let Name(Value)
		Call testsuitesElement.setAttribute("name", Value)
	End Property

	Public Property Get Name()
		Name = testsuitesElement.getAttribute("name")
	End Property

	'Public Method

	Public Function Xml()
		Xml = objDOM.Xml
	End Function
	

	Public Sub BeginTest(Name)
		Call testsuitesElement.setAttribute("name", Name)
		testsuitesTimer = Timer
	End Sub

	Public Sub EndTest()
		Call testsuitesElement.setAttribute("time", TimeTaken(testsuitesTimer))
		Call testsuitesElement.setAttribute("tests", testcaseCountGlobal)
		Call testsuitesElement.setAttribute("failures", failureCountGlobal)
	End Sub
	
	Public Sub BeginTestsuite(Name)
		Set testsuiteElement = objDOM.createElement("testsuite")
		testsuiteCount = testsuiteCount + 1
		Call testsuiteElement.setAttribute("id", testsuiteCount)
		Call testsuiteElement.setAttribute("name", Name)
		Call testsuitesElement.appendChild(testsuiteElement)
		testcaseCount = 0'initialize testcaseCount
		failureCount = 0'initialize failureCount
		errorCount = 0'initialize errorCount
		testsuiteTimer = Timer
	End Sub
	
	Public Sub EndTestsuite()
		testcaseCountGlobal = testcaseCountGlobal + testcaseCount
		failureCountGlobal = failureCountGlobal + failureCount
		errorCountGlobal = errorCountGlobal + errorCount
		Call testsuiteElement.setAttribute("time", TimeTaken(testsuiteTimer))
		Call testsuiteElement.setAttribute("tests", testcaseCount)
		Call testsuiteElement.setAttribute("failures", failureCount)
		Call testsuiteElement.setAttribute("errors", errorCount)
	End Sub

	Public Sub BeginTestcase(Name)
		Set testcaseElement = objDOM.createElement("testcase")
		testcaseCount = testcaseCount + 1
		Call testcaseElement.setAttribute("id", testcaseCount)
		Call testcaseElement.setAttribute("classname", testsuiteElement.getAttribute("name"))
		Call testcaseElement.setAttribute("name", Name)
		Call testsuiteElement.appendChild(testcaseElement)
		testcaseTimer = Timer
	End Sub

	Public Sub EndTestcase()
		Call testcaseElement.setAttribute("time", TimeTaken(testcaseTimer))
	End Sub

	Public Sub AssertEquals(Expected, Actual, Message)
		On Error Resume Next
		'Assertion
		result = Expected = Actual 
		'If Error
		If Err Then
			Call AppendError()
			Exit Sub
		End If
		On Error Goto 0
		If result Then
			'append nothing
		Else
			Call AppendFailure("AssertEquals", Message)
		End If
	End Sub

	Public Sub AssertTrue(Condition, Message)
		Dim result: result = false
		On Error Resume Next
		'Assertion
		result = Condition
		'If Error
		If Err Then
			Call AppendError()
			Exit Sub
		End If
		On Error Goto 0
		If result Then
			'append nothing
		Else
			Call AppendFailure("AssertTrue", Message)
		End If
	End Sub

	Public Sub AssertFalse(Condition, Message)
		Dim result: result = true
		On Error Resume Next
		'Assertion
		result = Condition
		'If Error
		If Err Then
			Call AppendError()
			Exit Sub
		End If
		On Error Goto 0
		If result Then
			Call AppendFailure("AssertTrue", Message)
		Else
			'append nothing
		End If
	End Sub

	Public Sub AssertNull(Actual, Message)
		On Error Resume Next
		'Assertion
		result = IsNull(Actual)
		'If Error
		If Err Then
			Call AppendError()
			Exit Sub
		End If
		On Error Goto 0
		If result Then
			'append nothing
		Else
			Call AppendFailure("AssertNull", Message)
		End If
	End Sub

	'Private Method

	Private Sub AppendError()
		Set errorElement = objDOM.createElement("error")
		Call errorElement.setAttribute("type", Err.Source)
		Call errorElement.setAttribute("message", Err.Description)
		testcaseElement.appendChild(errorElement)
		errorCount = errorCount + 1
	End Sub

	Private Sub AppendFailure(CallerName, Message)
		Set failureElement = objDOM.createElement("failure")
		Call failureElement.setAttribute("type", CallerName)
		Call failureElement.setAttribute("message", Message)
		testcaseElement.appendChild(failureElement)
		failureCount = failureCount + 1
	End Sub

	Private Function TimeTaken(startTimer)
		TimeTaken = (Timer * 1000 - startTimer * 1000) / 1000
	End Function
End Class
