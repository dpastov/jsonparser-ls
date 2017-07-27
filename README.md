# jsonparser-ls
JSON Parser written in LotusScript or VBA

Here is an example how to use JSONParser.

	Dim parser As JSONParser
	Dim jsonObj As JSONObject
	Dim jsonArr As JSONArray
	Dim jsonString As String
	
	Set parser = New JSONParser

	'object
	jsonString = |{"array":[1,  2  ,   300.56  ]  ,  "boolean":true,"null":null,"number":123,"object":{"a":"b","c":"d","arr":["12","23",34.56],"e":"f","ho":true},"string":"Hello World"}|
	Set jsonObj = parser.parse(jsonString)
	'test
	print jsonObj.GetItem("array").Items(2) '300.56
	print IsNull(jsonObj.GetItem("null")) 'true
	print jsonObj.GetItem("number") '123
	print jsonObj.GetItem("object").getItem("c") 'd
	print jsonObj.GetItem("object").getItem("ho") 'true
	print jsonObj.GetItem("object").getItem("arr").Items(2) '34.56
	
	'array
	jsonString = |[{a:1,b:true,_dd:null},12,"13",true,{}]|
	Set jsonArr = parser.parse(jsonString)
	'test
	print jsonArr.Items(0).getItem("b") 'true
	print jsonArr.Items(1) '12
	print jsonArr.Items(2) '13
	print jsonArr.Items(3) 'true
	print TypeName(jsonArr.Items(4)) '"JSONOBJECT"
