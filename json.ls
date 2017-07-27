Option Public
Option Declare

Class JSONParser
	Private m_length As Integer
	Private m_decimalSep As String
	
	Sub New()
	End Sub
	
	Private Property Get getDecimalSep As String
		Dim session As NotesSession
		Dim international As NotesInternational
		
		If m_decimalSep = "" Then
			Set session = New NotesSession()
			Set international = session.International
			m_decimalSep = international.DecimalSep
		End If
		
		getDecimalSep = m_decimalSep
	End Property
	
	private Property Get length As integer
		length = m_length
	End Property
	
	Function parse(jsonString As string) As variant
		Dim res As Variant
		Dim index1 As Integer
		Dim index2 As Integer
		
		m_length = Len(jsonString)
		
		index1 = InStr(jsonString, "{")
		index2 = InStr(jsonString, "[")
	
		If index1 > 0 And (index1 < index2 or index2 = 0) Then
			Set res = parseObject(jsonString, index1 + 1)
		ElseIf index2 > 0 And (index2 < index1 Or index1 = 0) Then
			Set res = parseArray(jsonString, index2 + 1)
		End If
		
		Set parse = res
	End Function
	
	Private Function parseObject(jsonString As String, index As integer) As JSONObject
		Dim res As JSONObject
		Dim propertyValue As Variant
		Dim propertyName As String
		Dim objectEnd As Integer
		Dim nextPair As integer

		Set res = New JSONObject()

		nextPair = InStr(index, jsonString, ":")
		objectEnd = InStr(index, jsonString, "}")
		While nextPair < objectEnd And nextPair > 0 And objectEnd > 0
			propertyName = findPropertyName(jsonString, index)
			index = InStr(index, jsonString, ":")
			index = index + 1
			
			Call renderValue(jsonString, index, propertyValue)
			Call res.AddItem(propertyName, propertyValue)
			
			nextPair = InStr(index, jsonString, ":")
			objectEnd = InStr(index, jsonString, "}")
		Wend
		
		index = objectEnd + 1
		
		Set parseObject = res
	End Function
	
	Private Function parseArray(jsonString As String, index As Integer) As JSONArray
		Dim res As JSONArray
		Dim propertyValue As Variant
		Dim arrEnd As Integer
		Dim nextVal As Integer

		Set res = New JSONArray()
		
		nextVal = InStr(index, jsonString, ",")
		arrEnd = InStr(index, jsonString, "]")
		While nextVal < arrEnd And nextVal > 0 And arrEnd > 0
			Call renderValue(jsonString, index, propertyValue)
			Call res.AddItem(propertyValue)
			
			nextVal = InStr(index, jsonString, ",")
			arrEnd = InStr(index, jsonString, "]")
		Wend
		
		index = arrEnd + 1
		
		Set parseArray = res
	End Function
	
	Private Function renderValue(jsonString As String, index As Integer, propertyValue As Variant) As Variant
		Dim char As String
		Dim i As Integer
		
		For i = index To length
			char = Mid(jsonString, i, 1)
			
			If char = {"} Then
				index = i
				propertyValue = findElementString(jsonString, index)
				i = length
			ElseIf char Like {#} Then
				index = i
				propertyValue = findElementNumber(jsonString, index)
				i = length
			ElseIf char Like {[tfn]} Then
				index = i
				propertyValue = findElementLiteral(jsonString, index)
				i = length
			ElseIf char = "{" Then
				index = i
				Set propertyValue = parseObject(jsonString, index)
				i = length
			ElseIf char = "[" Then
				index = i + 1
				Set propertyValue = parseArray(jsonString, index)
				i = length
			End If
		Next
	End Function
	
	Private Function findElementNumber(jsonString As string, index As Integer) As variant
		Dim res As variant
		Dim elementEnd As String
		Dim char As string
		Dim i As integer
		
		elementEnd = |, ]}|	'to catch: close bracket, comma, space or }
		For i = index To length
			char = Mid(jsonString, i, 1)
			
			If InStr(elementEnd, char) Then
				res = Mid(jsonString, index, i - index)
				index = i
				i = length
			End If
		Next
		
		If InStr(res, ".") And getDecimalSep()<>"." Then
			res = Replace(res, ".", getDecimalSep())
		End If
		
		findElementNumber = cdbl(res)
	End Function
	
	Private Function findElementLiteral(jsonString As String, index As Integer) As Variant
		Dim res As string
		Dim elementEnd As String
		Dim char As String
		Dim i As Integer
		
		elementEnd = |, ]}|	'to catch: close bracket, comma, space or }
		For i = index To length
			char = Mid(jsonString, i, 1)
			
			If InStr(elementEnd, char) Then
				res = Mid(jsonString, index, i - index)
				index = i
				i = length
			End If
		Next
		
		Select Case res:
			Case "null":
				findElementLiteral = null
			Case "true":
				findElementLiteral = true
			Case "false":
				findElementLiteral = false
		End Select
	End Function
	
	'find element in json string
	Private Function findElementString(jsonString As String, index As Integer) As String
		Dim res As String
		Dim index1 As Integer
		Dim index2 As Integer
		
		index1 = InStr(index, jsonString, {"})
		If index1 = 0 Then Exit Function
		
		index2 = InStr(index1 + 1, jsonString, {"})

		res = Mid(jsonString, index1 + 1, index2 - index1 - 1)
		
		index = index2 + 1
		
		findElementString = res
	End Function
	
	'find property name
	Private Function findPropertyName(jsonString As String, index As Integer) As String
		Dim res As String
		Dim propertyNameEnd As String
		Dim char As String
		Dim i As Integer
		
		'property start with character
		For i = index To length
			char = Mid(jsonString, i, 1)

			If char Like {[a-zA-Z_]} Then
				res = char
				index = i + 1
				i = length
			End If
		Next

		'rest of property could be characters and numbers etcx
		propertyNameEnd = | :"'|
		For i = index To length
			char = Mid(jsonString, i, 1)

			If InStr(propertyNameEnd, char) Then
				index = i
				i = length
			Else
				res = res + char
			End If
		Next
		
		findPropertyName = res
	End Function
End Class

Class JSONObject
	Private m_items List As Variant
	
	Sub New()
	End Sub
	
	Public Property Get Items As Variant
		Items = Me.m_items
	End Property
	
	Public Sub AddItem(itemName As String, itemVal As Variant)
		If IsObject(itemVal) Then
			Set me.m_items(itemName) = itemVal
		Else
			me.m_items(itemName) = itemVal
		End If
	End Sub
	
	Public Function GetItem(itemName As String) As Variant
		If IsElement(m_items(itemName)) Then
			If IsObject(me.Items(itemName)) Then
				Set GetItem = me.Items(itemName)
			Else
				GetItem = me.Items(itemName)
			End If
		Else
			GetItem = Null
		End If
	End Function
End Class

Class JSONArray
	Private m_items() As Variant
	Private m_size As integer

	Sub New()
	End Sub
	
	Public Property Get Items As Variant
		Items = Me.m_items
	End Property
	
	Public Property Get Size As integer
		Size = Me.m_size
	End Property
	
	Public Sub AddItem(itemVal As Variant)
		m_size = m_size + 1
		ReDim Preserve m_items(0 To m_size - 1) As Variant
				
		If IsObject(itemVal) Then
			Set m_items(Size - 1) = itemVal
		Else
			m_items(Size - 1) = itemVal
		End If
	End Sub
End Class