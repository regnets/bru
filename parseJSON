'AM_Funktion_parseJSON

Function printDictionary( objDic, level )
	Dim strResult, objItem, strSpacer, i
	If VarType( objDic ) = 9 Then
		strResult = strResult & "Dictionary auf Level " & level & " (Key: Value) " & Chr(10)
		For i = 0 To level 
			strSpacer = strSpacer & "_____"
		Next
		For Each objItem In objDic
			If VarType( objDic.Item(objItem) ) = 9 Then
				strResult = strResult & strSpacer & "'" & objItem & "': " & printDictionary( objDic.Item(objItem), (level+1) )
			Else
				strResult = strResult & strSpacer & "'" & objItem & "': '" & objDic.Item(objItem) & "'" & Chr(10)
			End If
		Next
	Else
		strResult = "Objekt ist kein Dictionary" & Chr(10)
	End If
	printDictionary = strResult
End Function

Function getDictionaryArray( strJSON )
	Dim objDic
	Set objDic = CreateObject("Scripting.Dictionary")
	Dim aPos, ePos 
	aPos = 1
	strJSON = Trim(Replace(strJSON, "\"&Chr(34), "\QUOTE"))
	Do While aPos < Len( strJSON )
		If Mid( strJSON, aPos, 1 ) = "[" Then
			ePos = findMatchPosition( strJSON, "[", "]", aPos)
			If Not ePos = 0 Then
				Set objDic = getDictionaryArray( Substring(strJSON, aPos+1, ePos-1) )
				aPos = ePos 'vorspulen
			Else
				'Schliessende eckige Klammer nicht gefunden
			End If
		ElseIf Mid( strJSON, aPos, 1 ) = "{" Then
			ePos = findMatchPosition( strJSON, "{", "}", aPos)
			If Not ePos = 0 Then
				Call objDic.Add( CStr(objDic.Count +1), getDictionary( Substring(strJSON, aPos, ePos) ))
				aPos = ePos 'vorspulen
			Else
				'Schliessende runde Klammer nicht gefunden
			End If
		End If
		aPos = aPos + 1	
	Loop
	
	Set getDictionaryArray = objDic
End Function

Function getDictionary( strJSON )
	Dim aPos, ePos, objDic, ePos2
	Set objDic = CreateObject("Scripting.Dictionary")

	aPos = 1
	ePos = 1
	Dim cur, isQuoted
	isQuoted = False
	If InStr(strJSON, ":")>0 Then
		Do While aPos < Len(strJSON) And ePos <= Len(strJSON)
			'Zeichen in Cursor lesen
			cur = Mid(strJSON, aPos, 1)
			
			'Zeichen ist ein oeffnendes Anfuehrungszeichen
			If cur = Chr(34) And isQuoted = False Then
				isQuoted = True
			'Zeichen ist ein schliessendes Anfuehrungszeichen
			ElseIf cur = Chr(34) And isQuoted = True Then
				isQuoted = False
			End If
			
			If cur = "}" And Not isQuoted Then
				Exit Do
			End If
			
			'Zeichen ist eine oeffnende Klammer {
			'Neues Dictionary erstellen und rekursiv aufrufen
			If cur = "{" And Not isQuoted Then
				Do While Not cur = "}" And aPos<Len(strJSON) Or isQuoted
					
					'Leerzeichen und Zeilenumbrueche ueberspringen
					Do While cur = " " Or cur = Chr(10) Or cur = Chr(13)
						aPos = aPos +1
						cur = Mid(strJSON, aPos, 1)
					Loop
				
					' aPos enthaelt nun oeffnende Klammer
					' suchen nach Doppelpunkt und Feldnamen ausschneiden
					ePos = InStr(aPos +1, strJSON, ":")
					
					'Rausspringen aus Schleife, wenn keine Feldnamen/ werte mehr
					'folgen
					If Not ePos > 0 Then
						Exit Do
					End If
					
					Dim strFeldname, strWert
					strFeldname = Substring( strJSON, aPos +1, ePos -1 )
					strFeldname = Replace(strFeldname, Chr(34), "")
					strFeldname = Replace(strFeldname, Chr(10), "")
					strFeldname = Replace(strFeldname, Chr(13), "")
					strFeldname = Trim(strFeldname)
									
					'Anfangsposition hinter Doppelpunkt setzen
					aPos = ePos + 1
					cur = Mid(strJSON, aPos, 1)
					
					'Leerzeichen und Zeilenumbrueche ueberspringen
					Do While cur = " " Or cur = Chr(10) Or cur = Chr(13)
						aPos = aPos +1
						cur = Mid(strJSON, aPos, 1)
					Loop
					
					
					Dim nKomma, nAnfuehrung, nKlammerZu, nEckKlammerAuf
					nKomma = InStr(aPos, strJSON, ",")
					If nKomma = 0 Then
						nKomma = Len(strJSON)+1
					End If
					
					nAnfuehrung = InStr(aPos, strJSON, Chr(34))
					If nAnfuehrung = 0 Then
						nAnfuehrung = Len(strJSON)+1
					End If
					
					nKlammerZu = InStr(aPos, strJSON, "}")
					If nKlammerZu = 0 Then
						nKlammerZu = Len(strJSON)+1
					End If
					
					nEckKlammerAuf = InStr(aPos, strJSON, "[")
					If nEckKlammerAuf = 0 Then
						nEckKlammerAuf = Len(strJSON)+1
					End If
					
					If nAnfuehrung < nKlammerZu And nKlammerZu <  Len(strJSON) Then
						'Naechste Klammer zu, wenn innerhalb von Anfuehrungszeichen
						nKlammerZu = InStr(nKlammerZu+1, strJSON, "}")
					End If

					If nKomma < nAnfuehrung And nKomma < nEckKlammerAuf And nKomma < nKlammerZu Then
						'Komma, Feldwert ohne Anfuehrungszeichen (z.B. einfache Zahl)
						strWert = Trim(Substring(strJSON, aPos, nKomma-1))
						aPos = nKomma
					ElseIf nKlammerZu < nAnfuehrung And nKlammerZu < nKomma And nKlammerZu < nEckKlammerAuf Then
						'schliessende geschweifte Klammer, Feldwert ist ein leeres JSON Array
						Set strWert = CreateObject("Scripting.Dictionary")
						aPos = nKlammerZu - 1
					ElseIf nKlammerZu < nKomma And nKlammerZu < nEckKlammerAuf And nKlammerZu < nAnfuehrung Then
						'schliessende geschweifte Klammer, Feldwert ist der letzte im Element
						strWert = Trim(Substring(strJSON, aPos, nKlammerZu-1))
						aPos = nKlammerZu-1
					ElseIf nAnfuehrung < nKomma And nKomma < nKlammerZu And nAnfuehrung < nEckKlammerAuf Then
						'Anfuehrungszeichen, Felderwert ist von Anfuehrungszeichen eingeschlossen
						ePos2 = InStr(aPos+1, strJSON, Chr(34))
						strWert = Substring(strJSON, aPos, ePos2)
						aPos = ePos2
					ElseIf nAnfuehrung < nKlammerZu And nKlammerZu < nKomma Then
						'Anfuehrungszeichen, Felderwert ist von Anfuehrungszeichen 	eingeschlossen und der letzte im Element
						strWert = Trim(Substring(strJSON, aPos, InStr(aPos+1, strJSON, Chr(34)) ))
						aPos = nKlammerZu-1
					ElseIf nEckKlammerAuf < nAnfuehrung And nEckKlammerAuf < nKlammerZu And nEckKlammerAuf < nKomma Then
						'Es folgt ein Array im Array
						ePos2 = findMatchPosition( strJSON, "[", "]", aPos)
						Set strWert = getDictionaryArray(Substring(strJSON, aPos, ePos2 ) )
						aPos = ePos2 -1 
					End If
					
					If VarType(strWert) = 9 Then
						'Dictionary

					Else
						strWert = Replace(strWert, Chr(34), "")
						strWert = Replace(strWert, "\QUOTE", Chr(34) )
					End If
					
					Call objDic.Add(strFeldname, strWert)
					
					aPos = aPos +1
					cur = Mid(strJSON, aPos, 1)
				Loop
			
			Else
				aPos = aPos +1
			End If
			aPos = aPos +1
		Loop
	End If
	Set getDictionary = objDic
End Function


Function findMatchPosition( strText, openChar, closeChar, aPos)
	Dim ePos, qFlag, nChar, x, y, nTextLength, cur
	nTextLength = Len(strText)
	'Werte kopieren, sonst Call by Reference
	x = aPos
	y = 0
	nChar = 0
	Do While aPos < nTextLength
		cur = Mid(strText, x , 1)
		If cur = Chr(34) And qFlag = True Then
			qFlag = False
		ElseIf cur = Chr(34) And qFlag = False Then
			qFlag = True
		ElseIf cur = openChar And qFlag = False Then
			nChar = nChar + 1
		ElseIf cur = closeChar And qFlag = False Then
			nChar = nChar - 1
			If nChar = 0 Then
				y = x 
				Exit Do
			End If
		End If
		x  = x  +1
	Loop
	findMatchPosition = y
End Function

Function Substring(strJSON, aPos, ePos)
	Substring = Mid(strJSON, aPos, ePos-aPos +1)
End Function
