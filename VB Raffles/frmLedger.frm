VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRaffle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E-Raffles"
   ClientHeight    =   11655
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   17205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11655
   ScaleWidth      =   17205
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   0
      Picture         =   "frmLedger.frx":0000
      ScaleHeight     =   2.385
      ScaleMode       =   5  'Inch
      ScaleWidth      =   11.885
      TabIndex        =   17
      Top             =   7920
      Width           =   17175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Caption         =   "Winners"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   4815
      Left            =   0
      TabIndex        =   2
      Top             =   4080
      Width           =   17175
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   16680
         Top             =   120
      End
      Begin MSComctlLib.ListView lstwinner 
         Height          =   3375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   16935
         _ExtentX        =   29871
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   64
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   17175
      Begin VB.TextBox txtprice 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   21
         Text            =   "Prize"
         Top             =   2880
         Width           =   3015
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   13440
         Top             =   3000
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   435
         Left            =   14280
         TabIndex        =   18
         Top             =   2880
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   767
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   5
         Value           =   5
         TextPosition    =   1
      End
      Begin VB.Timer timerLights 
         Enabled         =   0   'False
         Interval        =   230
         Left            =   6720
         Top             =   3000
      End
      Begin VB.Timer timerDuration 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   840
         Top             =   480
      End
      Begin VB.Timer timerExec 
         Enabled         =   0   'False
         Interval        =   75
         Left            =   240
         Top             =   480
      End
      Begin VB.CommandButton cmdSpin 
         Caption         =   "SPIN"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7320
         TabIndex        =   1
         Top             =   2880
         Width           =   2895
      End
      Begin MSComctlLib.ListView lstParticipants 
         Height          =   975
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblcountWinner 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Winners:"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label lblpartCount 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Count: "
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   14280
         TabIndex        =   23
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Prize"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Duration"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   20.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   14640
         TabIndex        =   20
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label lblMult 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   15000
         TabIndex        =   19
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblColor21 
         BackColor       =   &H000000C0&
         Height          =   615
         Left            =   13320
         TabIndex        =   14
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label lblColor22 
         BackColor       =   &H000000C0&
         Height          =   615
         Left            =   12120
         TabIndex        =   13
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label lblColor23 
         BackColor       =   &H000000C0&
         Height          =   615
         Left            =   10920
         TabIndex        =   12
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label lblColor13 
         BackColor       =   &H000000C0&
         Height          =   615
         Left            =   6000
         TabIndex        =   11
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label lblColor12 
         BackColor       =   &H000000C0&
         Height          =   615
         Left            =   4800
         TabIndex        =   10
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label lblColor11 
         BackColor       =   &H000000C0&
         Height          =   615
         Left            =   3600
         TabIndex        =   9
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label lblpid 
         Caption         =   "Label1"
         Height          =   495
         Left            =   1560
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblDateTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Date Time"
         BeginProperty Font 
            Name            =   "Bahnschrift"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   16935
      End
      Begin VB.Label lblDept 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Department - Designation"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   16935
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   24
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   16815
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status: Prototype; Version 0.8.5 (202312-19)"
      Height          =   255
      Left            =   10560
      TabIndex        =   16
      Top             =   11400
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by: Rex V.; Design By Kim O.; Unauthorized use is PROHIBITED"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   11400
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   12120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5625
   End
   Begin VB.Menu mnFile 
      Caption         =   "File"
      Begin VB.Menu mnImportPart 
         Caption         =   "Import Participants"
      End
      Begin VB.Menu mnExPart 
         Caption         =   "Export Participants"
      End
      Begin VB.Menu mnExport 
         Caption         =   "Export Winners"
      End
      Begin VB.Menu mnReset 
         Caption         =   "Reset Raffles"
      End
   End
End
Attribute VB_Name = "frmRaffle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSpin_Click()
timerLights.Enabled = True

timerDuration.Enabled = True
timerExec.Enabled = True
Me.cmdSpin.Enabled = False

'for Counter
Timer1.Enabled = True

'Time Correction
lblMult.Caption = lblMult.Caption - 1
End Sub

Private Sub Form_Click()
Unload frmSplash
End Sub

Private Sub Form_Load()
viewParticipants
viewWinners

partClount
WinnerCount
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Frame1_Click()
Unload frmSplash
End Sub

Private Sub Frame2_Click()
Unload frmSplash
End Sub


Private Sub Label5_Click()

End Sub



Private Sub lblDateTime_Click()
Unload frmSplash
End Sub

Private Sub lblDept_Click()
Unload frmSplash
End Sub

Private Sub lblName_Click()
Unload frmSplash
End Sub

Private Sub mnExPart_Click()
    On Error GoTo ErrorHandler

    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = DBstr()
    conn.Open

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    ' Execute the SELECT query
    rs.Open "SELECT pid, Name, Department, Designation, Organization, AddedBy " & _
            "FROM participants", conn, adOpenKeyset, adLockOptimistic, adCmdText

    ' Display Save As dialog to choose the save location
    Dim fileDialog As Object
    Set fileDialog = CreateObject("MSComDlg.CommonDialog")

    With fileDialog
        .DialogTitle = "Save CSV File"
        .Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        .FilterIndex = 1
        .ShowSave
        filePath = .FileName
    End With

    ' Check if the user canceled the Save As dialog
    If filePath = "" Then
        Exit Sub
    End If

    ' Open the CSV file for writing
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open filePath For Output As fileNumber

    ' Write the headers to the CSV file
    Print #fileNumber, "pid,Name,Department,Designation,Organization,AddedBy"

    ' Export data to CSV
    Do While Not rs.EOF
        Print #fileNumber, ReplaceNull(rs("pid")) & "," & _
                            EscapeCSV(ReplaceNull(rs("Name"))) & "," & _
                            EscapeCSV(ReplaceNull(rs("Department"))) & "," & _
                            EscapeCSV(ReplaceNull(rs("Designation"))) & "," & _
                            EscapeCSV(ReplaceNull(rs("Organization"))) & "," & _
                            EscapeCSV(ReplaceNull(rs("AddedBy")))

        rs.MoveNext
    Loop

    ' Close the CSV file and clean up
    Close fileNumber
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    MsgBox "Data exported to CSV successfully.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error Exporting Data: " & Err.Description, vbCritical, "Error"

End Sub

Private Sub mnExport_Click()
' Display Save As dialog to choose the save location
    Dim fileDialog As Object
    Set fileDialog = CreateObject("MSComDlg.CommonDialog")

    With fileDialog
        .DialogTitle = "Save CSV File"
        .Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        .FilterIndex = 1
        .ShowSave
        filePath = .FileName
    End With

    ' Check if the user canceled the Save As dialog
    If filePath = "" Then
        Exit Sub
    End If

    ' Open the CSV file for writing
    fileNumber = FreeFile
    Open filePath For Output As fileNumber

    ' Write the headers to the CSV file
    Print #fileNumber, "Name,Department-Designation,DateTime,Price"

    ' Export data to CSV
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Set conn = New ADODB.Connection
    conn.ConnectionString = DBstr()
    conn.Open

    Set rs = New ADODB.Recordset
    rs.Open "SELECT Name, Station AS Department_Designation, DateTime, Price " & _
            "FROM winner ORDER BY DateTime DESC", conn, adOpenKeyset, adLockOptimistic, adCmdText

    Do While Not rs.EOF
        i = i + 1
        ' Write data to the CSV file
        Print #fileNumber, EscapeCSV(rs("Name")) & "," & _
                        EscapeCSV(IIf(IsNull(rs("Department_Designation")), "", rs("Department_Designation"))) & "," & _
                        EscapeCSV(IIf(IsNull(rs("DateTime")), "", rs("DateTime"))) & "," & _
                        EscapeCSV(IIf(IsNull(rs("Price")), "", rs("Price")))
        rs.MoveNext
    Loop

    ' Close the CSV file
    Close fileNumber

    MsgBox "Data exported to CSV successfully.", vbInformation

    ' Clean up
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
End Sub

Private Sub mnImportPart_Click()
On Error GoTo ErrorHandler

Dim conn As ADODB.Connection
Set conn = New ADODB.Connection
conn.ConnectionString = DBstr()
conn.Open

' Prompt the user to select a CSV file
Dim fileDialog As Object
Set fileDialog = CreateObject("MSComDlg.CommonDialog")

With fileDialog
    .DialogTitle = "Select CSV File to Import"
    .Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    .FilterIndex = 1
    .ShowOpen
    filePath = .FileName
End With

' Check if the user canceled the Open dialog
If filePath = "" Then
    Exit Sub
End If

' Truncate the existing data in the participants table
conn.Execute "TRUNCATE TABLE participants"

' Open the CSV file for reading
Dim fileNumber As Integer
fileNumber = FreeFile
Open filePath For Input As fileNumber

' Skip the header line
Line Input #fileNumber, csvHeader

' Read and insert data from the CSV file
Do While Not EOF(fileNumber)
    Line Input #fileNumber, csvLine

    ' Use parameterized query to handle special characters
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandText = "INSERT INTO participants (pid, Name, Department, Designation, Organization, AddedBy) " & _
                      "VALUES (?, ?, ?, ?, ?, ?)"

    ' Split the CSV line into an array
    csvFields = Split(csvLine, ",")

    ' Set parameters
    Dim parameterValue As String
    For i = 0 To UBound(csvFields)
        parameterValue = Trim(csvFields(i))
        
        ' Check if the field is quoted and remove quotes
        If Left(parameterValue, 1) = """" And Right(parameterValue, 1) = """" Then
            parameterValue = Mid(parameterValue, 2, Len(parameterValue) - 2)
        End If
        
        ' Replace single quotes with double single quotes for SQL
        parameterValue = Replace(parameterValue, "'", "''")
        
        ' If the parameter is empty, replace with NULL
        If Len(parameterValue) = 0 Then
            parameterValue = "NULL"
        End If
        
        ' Add the parameter to the command
        cmd.Parameters.Append cmd.CreateParameter("param" & i, adVarChar, adParamInput, Len(parameterValue), parameterValue)
    Next i

    ' Execute the parameterized query
    cmd.Execute
Loop

' Close the CSV file and clean up
Close fileNumber
conn.Close
Set conn = Nothing

MsgBox "Participants imported successfully.", vbInformation

partClount
Exit Sub

ErrorHandler:
    MsgBox "Error Importing Participants: " & Err.Description, vbCritical, "Error"

    ' Log the error and details to a text file
    Dim logFileNumber As Integer
    logFileNumber = FreeFile
    Open "ImportErrorLog.txt" For Append As logFileNumber
    Print #logFileNumber, "Error Description: " & Err.Description
    Print #logFileNumber, "CSV Line: " & csvLine
    Print #logFileNumber, "Query: " & Replace(cmd.CommandText, "?", paramString, 1, -1, vbTextCompare)
    Print #logFileNumber, "------------------------"
    Close logFileNumber

    Resume Next

End Sub

Private Sub mnReset_Click()
 On Error GoTo ErrorHandler

    ' Prompt the user to confirm the truncation
    Dim userInput As String
    userInput = InputBox("Enter 'YES DELETE WINNERS' to proceed. This will be irreversible.")

    ' Check if the entered value is correct
    If UCase(userInput) = "YES DELETE WINNERS" Then
        ' User entered the correct value, proceed with truncation
        Dim conn As ADODB.Connection
        Set conn = New ADODB.Connection
        conn.ConnectionString = DBstr()
        conn.Open

        ' Execute the TRUNCATE TABLE statement
        conn.Execute "TRUNCATE TABLE winner"

        ' Clean up
        conn.Close
        Set conn = Nothing

        MsgBox "Winners table has been RESET.", vbInformation
        
        'Display Winners
        viewWinners
    Else
        ' User entered an incorrect value or canceled, do nothing
        MsgBox "Winners RESET canceled. Incorrect input or canceled by user.", vbInformation
    End If
    
    'Recount the Number of Entries After Reset
        WinnerCount
        partClount
    
    Exit Sub

ErrorHandler:
    MsgBox "Error Truncating Winner Table: " & Err.Description, vbCritical, "Error"

End Sub

Private Sub Slider1_Change()
timerDuration.Interval = Slider1.Value * 1000
lblMult.Caption = Slider1.Value
End Sub



Private Sub Slider1_KeyPress(KeyAscii As Integer)
Me.cmdSpin.SetFocus
End Sub

Private Sub Timer1_Timer()
If lblMult.Caption < 1 Then
    lblMult.Caption = 0
Else
lblMult.Caption = lblMult.Caption - 1
End If
End Sub

Private Sub timerDuration_Timer()

displayRandomEntry

'Function StopSpin Routine
StopSpin
frmSplash.Show

SaveWinnerData

'Return to old status
        ' Dark Red
        lblColor11.BackColor = RGB(139, 0, 0)
        lblColor12.BackColor = RGB(139, 0, 0)
        lblColor13.BackColor = RGB(139, 0, 0)

        lblColor21.BackColor = RGB(139, 0, 0)
        lblColor22.BackColor = RGB(139, 0, 0)
        lblColor23.BackColor = RGB(139, 0, 0)
        timerLights.Enabled = False
        
        'for Counter
        Timer1.Enabled = False
        
        WinnerCount
        partClount
End Sub

Private Sub timerExec_Timer()
'displayRandomEntry

' Check if both labels are "N/A" (i.e., no eligible participants)
    If frmRaffle.lblName.Caption = "N/A" And frmRaffle.lblDept.Caption = "N/A" Then
        ' Disable the timers when there are no eligible participants
        timerExec.Enabled = False
        timerDuration.Enabled = False

        ' Show alert if there are no eligible participants
        MsgBox "There are no Eligible Winners", vbInformation, "Error"

        ' Enable the Spin Button
        Me.cmdSpin.Enabled = True

        ' Exit the subroutine since there's no need to continue
        Exit Sub
    End If

    ' If the labels are not "N/A", generate and display random strings for the labels
    frmRaffle.lblName.Caption = generateRandomString(25) ' 10-character random string for Name
    frmRaffle.lblDept.Caption = generateRandomString(35) ' 15-character random string for Department

End Sub

Private Sub timerLights_Timer()
Static ColorToggle As Boolean ' Static variable retains its value between calls

    If ColorToggle Then
        ' Dark Red
        lblColor11.BackColor = RGB(255, 0, 0)
        lblColor12.BackColor = RGB(0, 0, 255)
        lblColor13.BackColor = RGB(255, 0, 0)

        lblColor21.BackColor = RGB(255, 0, 0)
        lblColor22.BackColor = RGB(0, 0, 255)
        lblColor23.BackColor = RGB(255, 0, 0)
    Else
        ' Dark Blue
        lblColor11.BackColor = RGB(0, 0, 255)
        lblColor12.BackColor = RGB(255, 0, 0)
        lblColor13.BackColor = RGB(0, 0, 255)

        lblColor21.BackColor = RGB(0, 0, 255)
        lblColor22.BackColor = RGB(255, 0, 0)
        lblColor23.BackColor = RGB(0, 0, 255)
    End If

    ' Toggle the color state for the next iteration
    ColorToggle = Not ColorToggle
End Sub
