# jsonparser-ls
JSON Parser written in LotusScript or VBA

Here is an example how to use JSONParser.


Function testJSONParse()
	On Error GoTo errh:

	Dim parser As JSONParser
	Dim jsonObj As JSONObject
	Dim jsonArr As JSONArray
	Dim jsonString As String
	Dim success As Integer
	Dim fail As integer
	
	Set parser = New JSONParser

	'object
	jsonString = |{"array":[1,  2  ,   300.56  ]  ,  "boolean":true,"null":null,"number":123,"object":{"a":"b","c":"d","arr":["12","23",34.56],"e":"f","ho":true},"string":"Hello World"}|
	Set jsonObj = parser.parse(jsonString)
	'test
	Call assertTrue(jsonObj.GetItem("array").Items(2)=300.56, success, fail)
	Call assertTrue(IsNull(jsonObj.GetItem("null")), success, fail)
	Call assertTrue(jsonObj.GetItem("number") = 123, success, fail)
	Call assertTrue(jsonObj.GetItem("object").getItem("c") = "d", success, fail)
	Call assertTrue(jsonObj.GetItem("object").getItem("ho") = true, success, fail)
	Call assertTrue(jsonObj.GetItem("object").getItem("arr").Items(2) = 34.56, success, fail)
	
	'array
	jsonString = |[{a:1,b:true,_dd:null},12,"12",true,{}]|
	Set jsonArr = parser.parse(jsonString)
	'test
	Call assertTrue(jsonArr.Items(0).getItem("b") = true, success, fail)
	Call assertTrue(jsonArr.Items(1) = 12, success, fail)
	Call assertTrue(jsonArr.Items(2) = "12", success, fail)
	Call assertTrue(jsonArr.Items(3) = true, success, fail)
	Call assertTrue(TypeName(jsonArr.Items(4)) = "JSONOBJECT", success, fail)
	
	If fail = 0 Then
		MsgBox "All " & CStr(success) & " tests have been passed", 64, "Success"
	Else
		MsgBox CStr(fail) & " tests have been failed", 16, "Fail"
	End If

	Exit Function
errh:
	MsgBox Error$ & " " & CStr(Erl)
	Exit Function
End Function

Public Function assertTrue(b As Boolean, success As integer, fail As integer)
	If b Then
		success = success + 1
	Else
		fail = fail + 1
	End If
End Function
