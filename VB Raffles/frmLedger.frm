VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRaffle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E-Raffles"
   ClientHeight    =   11655
   ClientLeft      =   45
   ClientTop       =   390
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
         Text            =   "Price"
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
         Top             =   3240
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
         Interval        =   10
         Left            =   240
         Top             =   480
      End
      Begin VB.CommandButton cmdSpin 
         Caption         =   "SPIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Price"
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
         Left            =   14520
         TabIndex        =   20
         Top             =   2760
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
         Left            =   14880
         TabIndex        =   19
         Top             =   1800
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
      Caption         =   "Status: Prototype; Version 0.8 (202312-10)"
      Height          =   255
      Left            =   10560
      TabIndex        =   16
      Top             =   11400
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by: Rex V.; Design By Kim Oclarit; Unauthorized use is PROHIBITED"
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
        
End Sub

Private Sub timerExec_Timer()
displayRandomEntry

If frmRaffle.lblName = "N/A" And frmRaffle.lblDept = "N/A" Then
'Disable The timers when there are no eligible participants"
timerExec.Enabled = False
timerDuration.Enabled = False

'Show alert if there no eligible participants
    MsgBox "There are no Eligible Winners", vbInformation, "Error"

'Enable the Spin Button
Me.cmdSpin.Enabled = True
End If

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
