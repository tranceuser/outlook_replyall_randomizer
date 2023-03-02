Option Explicit

Sub CabinResolved()
    Const SIGNATURE As String = "Technical Team"
    Const MIN_ROOM_NUMBER_DIGITS As Integer = 4
    
    Dim selectedMailItem As Object
    Dim greeting As String
    Dim timeOfDay As String
    Dim roomNumber As String
    Dim replies As Variant
    Dim randomIndex As Integer
    Dim reply As MailItem
    Dim regex As Object

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\d{" & MIN_ROOM_NUMBER_DIGITS & ",}"

    'On Error GoTo ErrorHandler

    If Application.ActiveExplorer.Selection.Count = 1 Then
        Set selectedMailItem = Application.ActiveExplorer.Selection(1)
        If TypeOf selectedMailItem Is MailItem Then
            roomNumber = ""
            If regex.Test(selectedMailItem.Subject) Then
                roomNumber = regex.Execute(selectedMailItem.Subject)(0)
            End If
            If roomNumber = "" Then
                roomNumber = "the mentioned stateroom"
            End If

            Select Case Hour(Now)
                Case Is < 12
                    timeOfDay = "morning"
                Case 12 To 17
                    timeOfDay = "afternoon"
                Case Else
                    timeOfDay = "evening"
            End Select

            greeting = "Good " & timeOfDay & ", " & vbNewLine & vbNewLine

            ' Define an array of reply messages
            replies = Array(greeting & "We wanted to let you know that the issue with cabin " & roomNumber & " has been resolved. If you have any further concerns, please let us know." & vbCrLf & vbCrLf & "Best Regards, " & vbCrLf & SIGNATURE, _
                               greeting & "We hope this message finds you well. We wanted to let you know that the issue with cabin " & roomNumber & " has been taken care of." & vbCrLf & vbCrLf & "Best Regards, " & vbCrLf & SIGNATURE, _
                               greeting & "We wanted to update you that the issue with cabin " & roomNumber & " has been resolved. If you need anything else, please don't hesitate to reach out." & vbCrLf & vbCrLf & "Best Regards, " & vbCrLf & SIGNATURE, _
                               greeting & "We wanted to let you know that the issue with cabin " & roomNumber & " has been addressed and resolved. Please let us know if you need anything else." & vbCrLf & vbCrLf & "Best Regards, " & vbCrLf & SIGNATURE, _
                               greeting & "We wanted to inform you that the issue with cabin " & roomNumber & " has been resolved. If there is anything else We can help with, please don't hesitate to ask." & vbCrLf & vbCrLf & "Best Regards, " & vbCrLf & SIGNATURE, _
                               greeting & "We wanted to let you know that your concerns regarding cabin " & roomNumber & " have been resolved. If there's anything else We can assist with, please let us know." & vbCrLf & vbCrLf & "Best Regards, " & vbCrLf & SIGNATURE)

            ' Randomly pick one of the replies
            randomIndex = Int((UBound(replies) - LBound(replies) + 1) * Rnd + LBound(replies))

            Set reply = selectedMailItem.ReplyAll
            reply.Body = replies(randomIndex)
            reply.Display
        End If
    End If
    
'ErrorHandler:
'    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Error"
'    Set regex = Nothing
End Sub
