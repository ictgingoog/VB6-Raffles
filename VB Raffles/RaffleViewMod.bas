Attribute VB_Name = "RaffleViewMod"
Function viewWinners() As String
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim Item As ListItem

    Set conn = New ADODB.Connection
    conn.ConnectionString = DBstr()
    conn.Open

    Set rs = New ADODB.Recordset

    rs.Open "SELECT wid, Name, Station, DateTime, Price " & _
        "FROM winner ORDER BY DateTime DESC", conn, adOpenKeyset, adLockOptimistic, adCmdText


    ' Clear the listview
    frmRaffle.lstwinner.ListItems.Clear
    frmRaffle.lstwinner.ColumnHeaders.Clear

    ' Specify the column headers
    frmRaffle.lstwinner.ColumnHeaders.Add , , "Name"
    frmRaffle.lstwinner.ColumnHeaders.Item(1).Width = 6500 ' Set width for the first column
    frmRaffle.lstwinner.ColumnHeaders.Add , , "Department- Designation"
    frmRaffle.lstwinner.ColumnHeaders.Item(2).Width = 6500
    frmRaffle.lstwinner.ColumnHeaders.Add , , "DateTime"
    frmRaffle.lstwinner.ColumnHeaders.Item(3).Width = 2700
    frmRaffle.lstwinner.ColumnHeaders.Add , , "Price"
    frmRaffle.lstwinner.ColumnHeaders.Item(4).Width = 1300
    
    ' Loop through the recordset and add each row to the listview
    Do While Not rs.EOF
        i = i + 1
        Set Item = frmRaffle.lstwinner.ListItems.Add(i, , rs("Name"))
        Item.SubItems(1) = IIf(IsNull(rs("Station")), "", rs("Station"))
        Item.SubItems(2) = IIf(IsNull(rs("DateTime")), "", rs("DateTime"))
        Item.SubItems(3) = IIf(IsNull(rs("Price")), "", rs("Price"))

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
End Function



Function viewParticipants() As String

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim Item As ListItem

    Set conn = New ADODB.Connection
    conn.ConnectionString = DBstr()
    conn.Open

    Set rs = New ADODB.Recordset

    rs.Open "SELECT pid, Name, Department, Designation, Organization, AddedBy " & _
            "FROM participants", conn, adOpenKeyset, adLockOptimistic, adCmdText

    frmRaffle.lstParticipants.ListItems.Clear
    frmRaffle.lstParticipants.ColumnHeaders.Clear

    frmRaffle.lstParticipants.ColumnHeaders.Add , , "pid"
    frmRaffle.lstParticipants.ColumnHeaders.Add , , "Name"
    frmRaffle.lstParticipants.ColumnHeaders.Add , , "Department"
    frmRaffle.lstParticipants.ColumnHeaders.Add , , "Designation"
    frmRaffle.lstParticipants.ColumnHeaders.Add , , "Organization"
    frmRaffle.lstParticipants.ColumnHeaders.Add , , "AddedBy"

    Do While Not rs.EOF
        i = i + 1
        Set Item = frmRaffle.lstParticipants.ListItems.Add(i, , rs("pid"))
        Item.SubItems(1) = IIf(IsNull(rs("Name")), "", rs("Name"))
        Item.SubItems(2) = IIf(IsNull(rs("Department")), "", rs("Department"))
        Item.SubItems(3) = IIf(IsNull(rs("Designation")), "", rs("Designation"))
        Item.SubItems(4) = IIf(IsNull(rs("Organization")), "", rs("Organization"))
        Item.SubItems(5) = IIf(IsNull(rs("AddedBy")), "", rs("AddedBy"))

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    

End Function
Function displayRandomEntry()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim randompid As String
    Dim randomName As String
    Dim randomDept As String
    Dim currentTime As String

    Set conn = New ADODB.Connection
    conn.ConnectionString = DBstr()
    conn.Open

    Set rs = New ADODB.Recordset

    ' Retrieve a random entry excluding pids from the winner table
    rs.Open "SELECT pid, Name, Department, Designation " & _
        "FROM participants " & _
        "WHERE NOT EXISTS " & _
        "(SELECT 1 FROM winner WHERE winner.pid = participants.pid) " & _
        "ORDER BY RAND() LIMIT 1", _
        conn, adOpenKeyset, adLockOptimistic, adCmdText

    If Not rs.EOF Then
        randompid = rs("pid").Value
        randomName = rs("Name").Value
        randomDept = rs("Department").Value & " - " & rs("Designation").Value
    Else
        'Set Countdown to Slider Value
        frmRaffle.lblMult.Caption = frmRaffle.Slider1
        
        StopSpin
        
        MsgBox "All participants have won!", vbInformation, "Raffle Over"
        'Execute SpotSpin Routine
        
        Exit Function
    End If

    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing

    ' Display the data in the form's labels
    frmRaffle.lblpid.Caption = randompid
    frmRaffle.lblName.Caption = randomName
    frmRaffle.lblDept.Caption = randomDept

    ' Generate and display the current timestamp
    currentTime = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    frmRaffle.lblDateTime.Caption = currentTime
End Function


Public Sub SaveWinnerData()
    On Error GoTo ErrorHandler ' Commented out for debugging purposes
    Dim conn As ADODB.Connection
    Dim sql As String

    ' Assuming you have a function named DBstr() that returns the connection string
    Set conn = New ADODB.Connection
    conn.ConnectionString = DBstr()
    conn.Open

    ' Assuming your labels are named lblpid, lblName, lblDept, and lblDateTime
    sql = "INSERT INTO winner (pid, Name, Station, Price, DateTime, batchID) VALUES (" & _
          "'" & EscapeInput(frmRaffle.lblpid.Caption) & "', " & _
          "'" & EscapeInput(frmRaffle.lblName.Caption) & "', " & _
          "'" & EscapeInput(frmRaffle.lblDept.Caption) & "', " & _
          "'" & EscapeInput(frmRaffle.txtprice.Text) & "', " & _
          "'" & EscapeInput(frmRaffle.lblDateTime.Caption) & "', 1)"

    ' Execute the SQL query to insert the values into the "winner" table
    Debug.Print sql ' Print the SQL statement for debugging purposes
    conn.Execute sql

    ' Clean up
    conn.Close
    Set conn = Nothing

    Exit Sub ' This line ensures that the program continues if there is no error

ErrorHandler:
    MsgBox "Error Saving Winner: " & Err.Description, vbCritical, "Error"
    End ' Terminate the program in case of an error
End Sub


