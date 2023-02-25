Option Explicit

dim Debugger, varStr, varInt, varBool, varNull
Set Debugger = New DevDebugger
	
varStr = "Test String"
varInt = 3
varBool = False
varNull = Null

Debugger.Print(varStr)
Debugger.Print(varInt)
Debugger.PrintName varBool, "TestVar"
Debugger.Print(varNull)

Class DevDebugger
	'Author: Nitemare
	'Date: 2023/02/24
	Private b, f, r, n, t
	
	Private Sub Class_Initialize
		b = ChrW( 8 )
		f = vbFormFeed
		r = vbCr
		n = vbLf
		t = vbTab
	End Sub
	
	Sub Print(ByVal objInput)
		wscript.echo "Print_Var => " & ParseVariable(objInput, 0) & vbNewLine
	End Sub
	
	Sub PrintName(ByVal objInput, inputName)
		wscript.echo "Print_Var '" & inputName & "' => " & ParseVariable(objInput, 0) & vbNewLine
	End Sub

	Private Function ParseVariable(ByVal varInput, Level)
		Dim valCount: valCount = Null
		dim valType, returnable, strArrayEntry, strFull, i, j, c
		
		Select Case VarType (varInput)
			Case vbNull
				valType = "Null"
				returnable = "null"
			Case vbBoolean
				valType = "Boolean"
				If varInput Then
					returnable = "true"
				Else
					returnable = "false"
				End If
			Case vbInteger , vbLong , vbSingle , vbDouble
				Select Case VarType (varInput)
					Case vbInteger
						valType = "Integer"
					Case vbLong
						valType = "Long"
					Case vbSingle
						valType = "Single"
					Case vbDouble
						valType = "Double"
				End Select
				returnable = varInput
			Case vbString
				valType = "String"
				valCount = Len (varInput)
				strFull = ""
				For i = 1 To Len (varInput)
					c = Mid (varInput, i, 1 )
					Select Case c
						Case """" strFull = strFull & "\"""
						Case "\" strFull = strFull & "\\"
						Case "/" strFull = strFull & "/"
						Case b strFull = strFull & "\b"
						Case f strFull = strFull & "\f"
						Case r strFull = strFull & "\r"
						Case n strFull = strFull & "\n"
						Case t strFull = strFull & "\t"
						Case Else
							If AscW(c) >= 0 And AscW(c) <= 31 Then
								c = Right ( "0" & Hex (AscW(c)), 2 )
								strFull = strFull & "\u00" & c
							Else
								strFull = strFull &  c
							End If
					End Select
				Next
				returnable = """" & strFull & """"
			Case vbArray + vbVariant
				If getDim(VarInput) <= 1 Then
					valType = "Array"
					valCount = 0
					For Each strArrayEntry In varInput
						Returnable = Returnable & vbNewLine & string(Level + 1, vbTab) & valCount & " => " & ParseVariable(strArrayEntry, Level + 1) 
						valCount = valCount + 1
					Next
				Else
					valType = "RectArray"
					valCount = Ubound(varInput,1) & "," & Ubound(varInput,2)
					ReDim arrvarInput(UBound(varInput, 1))
					 For i = 0 To UBound(varInput, 1)
						 ReDim arrProp(UBound(varInput, 2))
						 For j = 0 To UBound(varInput, 2)
							Returnable = Returnable & vbNewLine & string(Level + 1, vbTab) & "(" & i & "," & j & ")" & " => " & ParseVariable(varInput(i, j), Level + 1) 
						 Next
						 Returnable = Returnable & vbNewLine
					 Next
				End If
			Case vbObject
				valType = "Object"
				returnable = "<Objects Not Yet Supported>"
			Case Else
				valType = "Unkown"
				returnable = "<Unkown Variable Type: " & varType(obj) & " >"
		end select
		If Not IsNull(valCount) Then
			ParseVariable = valType & "(" & valCount & ") " & Returnable
		Else
			ParseVariable = valType & " " & Returnable
		end if
	End Function
	
	Private Function getDim(var)
		On Error Resume Next
		Dim tmp, i: i = 0
		Do While True
			i = i + 1
			tmp = UBound(var, i)
			If Err.Number <> 0 Then
				On Error GoTo 0
				getDim = i - 1
				exit function
			End If
		Loop
	End Function
End Class
