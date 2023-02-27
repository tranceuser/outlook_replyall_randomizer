Sub ReplyWithRandomGreetingAndResolvedIssue()

	Dim objItem As Object
	Dim strGreeting As String
	Dim strTimeOfDay As String
	Dim strRoomNumber As String
	Dim arrReplies As Variant
	Dim intRandomIndex As Integer
	Dim objReply As MailItem
	Dim regex As Object

	Set regex = CreateObject("VBScript.RegExp")
	regex.Pattern = "\d+"

	If Application.ActiveExplorer.Selection.Count = 1 Then
		Set objItem = Application.ActiveExplorer.Selection(1)
		If TypeOf objItem Is MailItem Then
			Set objRegEx = CreateObject("VBScript.RegExp")
			With objRegEx
				.Global = True
				.MultiLine = True
				.Pattern = "\b\d{4,}\b"
				strRoomNumber = .Execute(objItem.Subject)(0)
			End With

			If Not IsNumeric(strRoomNumber) Then
				strRoomNumber = "the mentioned stateroom" ' If the room number is not numeric, set strRoomNumber to a default value
			End If

			Select Case Hour(Now)
				Case Is < 12
					strTimeOfDay = "morning"
				Case 12 To 17
					strTimeOfDay = "noon"
				Case Else
					strTimeOfDay = "evening"
			End Select	

			strGreeting = "Good " & strTimeOfDay & ", " & vbCrLf & vbCrLf

            ' Define an array of reply messages
            arrReplies = Array(strGreeting & "We wanted to let you know that the issue with cabin " & strRoomNumber & " has been resolved. If you have any further concerns, please let us know." & vbCrLf & vbCrLf & "Best Regards, " & vbCrLf & "Technical Team", _
                               strGreeting & "We hope this message finds you well. We wanted to let you know that the issue with cabin " & strRoomNumber & " has been taken care of." & vbCrLf & vbCrLf & "Best Regards, " & vbCrLf & "Technical Team", _
                               strGreeting & "We wanted to update you that the issue with cabin " & strRoomNumber & " has been resolved. If you need anything else, please don't hesitate to reach out." & vbCrLf & vbCrLf & "Best Regards, " & vbCrLf & "Technical Team", _
                               strGreeting & "We wanted to let you know that the issue with cabin " & strRoomNumber & " has been addressed and resolved. Please let us know if you need anything else." & vbCrLf & vbCrLf & "Best Regards, " & vbCrLf & "Technical Team", _
                               strGreeting & "We wanted to inform you that the issue with cabin " & strRoomNumber & " has been resolved. If there is anything else We can help with, please don't hesitate to ask." & vbCrLf & vbCrLf & "Best Regards, " & vbCrLf & "Technical Team", _
                               strGreeting & "We wanted to let you know that your concerns regarding cabin " & strRoomNumber & " have been resolved. If     there's anything else We can assist with, please let us know." & vbCrLf & vbCrLf & "Best Regards, " & vbCrLf & "Technical Team")

            ' Randomly pick one of the replies
            intRandomIndex = Int((UBound(arrReplies) - LBound(arrReplies) + 1) * Rnd + LBound(arrReplies))

            Set objReply = objItem.ReplyAll
            objReply.Body = arrReplies(intRandomIndex)
            objReply.Display
        End If
    End If

End Sub
