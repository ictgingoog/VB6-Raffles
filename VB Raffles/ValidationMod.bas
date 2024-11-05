Attribute VB_Name = "ValidationMod"
Function EscapeInput(inputString As String) As String
    ' Replace single quotes with two single quotes
    Dim outputString As String
    outputString = Replace(inputString, "'", "''")

    ' Replace backslashes with two backslashes
    outputString = Replace(outputString, "\", "\\")

    ' Allow only numbers, letters (both uppercase and lowercase), space, period, hyphen, forward slash, colon, and percentage symbol
    Dim allowedChars As String
    allowedChars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz. -/:%"

    Dim i As Integer
    For i = 1 To Len(outputString)
        If InStr(allowedChars, Mid(outputString, i, 1)) > 0 Then
            ' Include the character if it's allowed
            EscapeInput = EscapeInput & Mid(outputString, i, 1)
        End If
    Next i
End Function


Function StopSpin()
    frmRaffle.timerExec.Enabled = False
    frmRaffle.timerDuration.Enabled = False


    frmSplash.lblName = frmRaffle.lblName
    frmSplash.lblDept = frmRaffle.lblDept
    frmSplash.lblPrice = "- " & frmRaffle.txtprice & " -"
    
    'Return to old status
        ' Dark Red
        frmRaffle.lblColor11.BackColor = RGB(139, 0, 0)
        frmRaffle.lblColor12.BackColor = RGB(139, 0, 0)
        frmRaffle.lblColor13.BackColor = RGB(139, 0, 0)

        frmRaffle.lblColor21.BackColor = RGB(139, 0, 0)
        frmRaffle.lblColor22.BackColor = RGB(139, 0, 0)
        frmRaffle.lblColor23.BackColor = RGB(139, 0, 0)
        frmRaffle.timerLights.Enabled = False
        
        'for Counter
        frmRaffle.Timer1.Enabled = False
        
        'Enable Spin button
        viewWinners
        frmRaffle.cmdSpin.Enabled = True
        
        'Set Countdown to Slider Value
        frmRaffle.lblMult.Caption = frmRaffle.Slider1

        frmRaffle.cmdSpin.SetFocus
End Function
