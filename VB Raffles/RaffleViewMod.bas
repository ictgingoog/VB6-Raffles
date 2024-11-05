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
    On Error GoTo ErrorHandler ' Add error handling

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim randompid As String
    Dim randomName As String
    Dim randomDept As String
    Dim currentTime As String
    Dim eligibleDepartments As Collection
    Dim departmentParticipants As Object ' Use Object for dictionary-like behavior
    Dim selectedDept As String
    Dim attempts As Long
    Dim randomDeptIndex As Long
    Dim randomParticipantIndex As Long
    Const MAX_ATTEMPTS As Long = 10 ' Limit the number of attempts to avoid infinite loops

    ' Initialize and open the connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = DBstr()
    conn.Open

    ' Initialize the recordset
    Set rs = New ADODB.Recordset

    ' Step 1: Retrieve all eligible departments and participants in a single query
    rs.Open "SELECT Department, pid, Name, Designation " & _
            "FROM participants " & _
            "WHERE NOT EXISTS " & _
            "(SELECT 1 FROM winner WHERE winner.pid = participants.pid)", _
            conn, adOpenStatic, adLockReadOnly, adCmdText

    ' Initialize collections to store eligible departments and participants
    Set eligibleDepartments = New Collection
    Set departmentParticipants = CreateObject("Scripting.Dictionary") ' Dictionary to store participants by department

    ' Process the result set
    Do While Not rs.EOF
        ' Group participants by department
        If Not departmentParticipants.Exists(rs("Department").Value) Then
            eligibleDepartments.Add rs("Department").Value
            departmentParticipants.Add rs("Department").Value, New Collection
        End If
        departmentParticipants(rs("Department").Value).Add Array(rs("pid").Value, rs("Name").Value, rs("Designation").Value)
        rs.MoveNext
    Loop
    rs.Close

    ' If no eligible departments are found, handle it
    If eligibleDepartments.Count = 0 Then
        frmRaffle.lblMult.Caption = frmRaffle.Slider1 ' Set Countdown to Slider Value
        StopSpin ' Stop the spinning/raffle process
        MsgBox "All participants have won!", vbInformation, "Raffle Over"
        Exit Function
    End If

    ' Step 2: Loop to find a department with eligible participants
    attempts = 0
    Do
        ' Generate a random index to select a department
        Randomize ' Initialize the random number generator
        randomDeptIndex = Int(Rnd() * eligibleDepartments.Count) + 1
        selectedDept = eligibleDepartments(randomDeptIndex)

        ' Check if the selected department has eligible participants
        If departmentParticipants(selectedDept).Count > 0 Then
            Exit Do ' Found a department with eligible participants
        End If

        ' Increment the attempt counter to avoid infinite loops
        attempts = attempts + 1
        If attempts >= MAX_ATTEMPTS Then
            MsgBox "Unable to find a department with eligible participants.", vbExclamation, "Error"
            Exit Function
        End If

    Loop

    ' Step 3: Randomly select a participant from the selected department
    randomParticipantIndex = Int(Rnd() * departmentParticipants(selectedDept).Count) + 1
    Dim selectedParticipant As Variant
    selectedParticipant = departmentParticipants(selectedDept)(randomParticipantIndex)

    ' Extract the random participant's details
    randompid = selectedParticipant(0) ' pid
    randomName = selectedParticipant(1) ' Name
    randomDept = selectedDept & " - " & selectedParticipant(2) ' Department - Designation

    ' Close the connection
    conn.Close
    Set conn = Nothing

    ' Display the data in the form's labels
    frmRaffle.lblpid.Caption = randompid
    frmRaffle.lblName.Caption = randomName
    frmRaffle.lblDept.Caption = randomDept

    ' Generate and display the current timestamp
    currentTime = Format(Now(), "yyyy-mm-dd hh:nn:ss") ' Corrected time format
    frmRaffle.lblDateTime.Caption = currentTime

    Exit Function

ErrorHandler:
    ' Handle any errors that occur
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
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

Function partClount()
On Error Resume Next

    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = DBstr()
    conn.Open

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    ' Execute the SELECT query with COUNT function
    rs.Open "SELECT COUNT(*) AS RecordCount FROM participants", conn, adOpenKeyset, adLockOptimistic, adCmdText

    ' Check if the recordset is not empty
    If Not rs.EOF Then
        ' Display the total number of entries in a label
        frmRaffle.lblpartCount.Caption = "Total Entries: " & rs("RecordCount")
    Else
        ' If the recordset is empty, display a message in the label
        frmRaffle.lblpartCount.Caption = "No Entries"
    End If

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    ' Resume execution after an error
    Resume Next
End Function

Function WinnerCount()
    On Error Resume Next

    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = DBstr()
    conn.Open

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    ' Execute the SELECT query with COUNT function for the "winner" table
    rs.Open "SELECT COUNT(*) AS RecordCount FROM winner", conn, adOpenKeyset, adLockOptimistic, adCmdText

    ' Check if the recordset is not empty
    If Not rs.EOF Then
        ' Display the total number of entries in a label
        ' Replace "LabelName" with the actual name of your label
        frmRaffle.lblcountWinner.Caption = "Total Winners: " & rs("RecordCount")
    Else
        ' If the recordset is empty, display a message in the label
        ' Replace "LabelName" with the actual name of your label
        frmRaffle.lblcountWinner.Caption = "No Winners Yet"
    End If

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    ' Resume execution after an error
    Resume Next
End Function

Function generateRandomString(length As Integer) As String
    Dim chars As String
    Dim result As String
    Dim i As Integer
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789" ' Characters to use
    result = ""
    
    ' Seed the random number generator
    Randomize
    
    ' Generate random characters
    For i = 1 To length
        result = result & Mid(chars, Int(Rnd() * Len(chars)) + 1, 1)
    Next i
    
    generateRandomString = result
End Function

