<script language="vbscript" runat="server" src="Class/VbsUnitTest.vbs"></script>
<script language="vbscript" runat="server" src="Calc.vbs"></script>
<%
    Response.AddHeader "Content-Type", "application/xml; charset=utf-8"
     Dim Expected, Actual
    Set Test = new VbsUnitTest

    'Begin Test
    Test.BeginTest Request.ServerVariables("SCRIPT_NAME")

    'Begin Testsuite
    Test.BeginTestsuite "Sum"

    'Begin Testcase
    Test.BeginTestcase "Sum(1, 2) returns 3"
    Expected = 3
    Actual = Sum(1, 2)
    Test.AssertEquals Expected, Actual, "Expected is '" & Expected & "'.but Actual is '" & Actual & "'"
    Test.EndTestcase
    'End Testcase

    'Begin Testcase
    Test.BeginTestcase "Sum(""1"", ""2"") returns 3"
    Expected = 3
    Actual = Sum("1", "2")
    Test.AssertEquals Expected, Actual, "Expected is '" & Expected & "'.but Actual is '" & Actual & "'"
    Test.EndTestcase
    'End Testcase
    
    'Begin Testcase
    Test.BeginTestcase "Sum(null, 2) returns 2"
    Expected = 2
    Actual = Sum(null, 2)
    Test.AssertEquals Expected, Actual, "Expected is '" & Expected & "'.but Actual is '" & Actual & "'"
    Test.EndTestcase
    'End Testcase
    
    'Begin Testcase
    Test.BeginTestcase "Sum(""abc"", 2) returns 2"
    Expected = 2
    Actual = Sum("abc", 2)
    Test.AssertEquals Expected, Actual, "Expected is '" & Expected & "'.but Actual is '" & Actual & "'"
    Test.EndTestcase
    
    Test.EndTestsuite
    'End Testsuite

    'Begin Testsuite
    Test.BeginTestsuite "Diff"
    
    'Begin Testcase
    Test.BeginTestcase "Diff(1, 2) returns -1"
    Expected = -1
    Actual = Diff(1, 2)
    Test.AssertEquals Expected, Actual, "Expected is '" & Expected & "'.but Actual is '" & Actual & "'"
    Test.EndTestcase
    'End Testcase

    'Begin Testcase
    Test.BeginTestcase "Diff(""1"", ""2"") returns -1"
    Expected = -1
    Actual = Diff("1", "2")
    Test.AssertEquals Expected, Actual, "Expected is '" & Expected & "'.but Actual is '" & Actual & "'"
    Test.EndTestcase
    'End Testcase

    'Begin Testcase
    Test.BeginTestcase "Diff(null, 2) returns -2"
    Expected = -2
    Actual = Diff(null, 2)
    Test.AssertEquals Expected, Actual, "Expected is '" & Expected & "'.but Actual is '" & Actual & "'"
    Test.EndTestcase
    'End Testcase

    'Begin Testcase
    Test.BeginTestcase "Diff(""abc"", 2) returns -2"
    Expected = -2
    Actual = Diff("abc", 2)
    Test.AssertEquals Expected, Actual, "Expected is '" & Expected & "'.but Actual is '" & Actual & "'"
    Test.EndTestcase
    'End Testcase
    
    Test.EndTestsuite
    'End Testsuite

    'Begin Testsuite
    Test.BeginTestsuite "IsOdd"

    'Begin Testcase
    Test.BeginTestcase "IsOdd(1) returns True"
    Actual = IsOdd(1)
    Test.AssertTrue Actual, "Expected is '" & Expected & "'.but Actual is '" & Actual & "'"
    Test.EndTestcase
    'End Testcase

    'Begin Testcase
    Test.BeginTestcase "IsOdd(2) returns False"
    Actual = IsOdd(2)
    Test.AssertFalse Actual, "Expected is '" & Expected & "'.but Actual is '" & Actual & "'"
    Test.EndTestcase
    'End Testcase
    
    Test.EndTestsuite
    'End Testsuite

    'Begin Testsuite
    Test.BeginTestsuite "IsEven"
    
    'Begin Testcase
    Test.BeginTestcase "IsEven(1) returns False"
    Actual = IsEven(1)
    Test.AssertFalse Actual, "Expected is '" & Expected & "'.but Actual is '" & Actual & "'"
    Test.EndTestcase
    'End Testcase

    'Begin Testcase
    Test.BeginTestcase "IsEven(2) returns True"
    Actual = IsEven(2)
    Test.AssertTrue Actual, "Expected is '" & Expected & "'.but Actual is '" & Actual & "'"
    Test.EndTestcase
    'End Testcase

    Test.EndTestsuite
    'End Testsuite

    Test.EndTest
    'End Test

    Response.Write(Test.Xml())
%>