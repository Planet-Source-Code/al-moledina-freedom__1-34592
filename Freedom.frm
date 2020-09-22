VERSION 5.00
Begin VB.Form frmFreedom 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6060
   ControlBox      =   0   'False
   Icon            =   "Freedom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   60
      TabIndex        =   7
      Top             =   -15
      Width           =   5940
      Begin VB.CommandButton cmdConfig 
         Caption         =   "C&onfigure"
         Height          =   375
         Left            =   180
         TabIndex        =   9
         Top             =   1500
         Width           =   1155
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Default         =   -1  'True
         Height          =   375
         Left            =   4800
         TabIndex        =   8
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Caption         =   "lblInfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   555
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   5715
      End
      Begin VB.Image Image1 
         Height          =   540
         Left            =   1740
         Picture         =   "Freedom.frx":08CA
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Image imgDay 
         Height          =   480
         Index           =   0
         Left            =   3180
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image imgDay 
         Height          =   480
         Index           =   1
         Left            =   3660
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image imgDay 
         Height          =   480
         Index           =   2
         Left            =   4140
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image imgCongrats 
         Height          =   690
         Left            =   60
         Picture         =   "Freedom.frx":0ECE
         Top             =   60
         Visible         =   0   'False
         Width           =   5820
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Today is.."
         Height          =   195
         Left            =   2220
         TabIndex        =   11
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lblMe 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Freedom® - brought to you absolutely FREE!! -  by ex-smoker Al Moledina.  Comments? Email me at amoledin@telusplanet.net"
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   300
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   5415
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   120
      Top             =   3780
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Top             =   3300
      Width           =   675
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   4080
      MaxLength       =   2
      TabIndex        =   1
      Top             =   3000
      Width           =   675
   End
   Begin VB.ComboBox cboStartDate 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4980
      TabIndex        =   3
      Top             =   3180
      Width           =   915
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   240
      Picture         =   "Freedom.frx":195F
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   9
      Left            =   5460
      Picture         =   "Freedom.frx":2229
      Top             =   3780
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   8
      Left            =   4920
      Picture         =   "Freedom.frx":2AF3
      Top             =   3780
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   7
      Left            =   4380
      Picture         =   "Freedom.frx":33BD
      Top             =   3780
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   6
      Left            =   3840
      Picture         =   "Freedom.frx":3C87
      Top             =   3780
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   5
      Left            =   3300
      Picture         =   "Freedom.frx":4551
      Top             =   3780
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   4
      Left            =   2760
      Picture         =   "Freedom.frx":4E1B
      Top             =   3780
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   3
      Left            =   2220
      Picture         =   "Freedom.frx":56E5
      Top             =   3780
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   2
      Left            =   1680
      Picture         =   "Freedom.frx":5FAF
      Top             =   3780
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   1
      Left            =   1140
      Picture         =   "Freedom.frx":6879
      Top             =   3780
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   0
      Left            =   720
      Picture         =   "Freedom.frx":7143
      Top             =   3780
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Price for a pack of 25 (dollars and cents, 00.00)"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   3300
      Width           =   3795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Number of cigarettes smoked per day"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   3000
      Width           =   3795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Quit Date"
      Height          =   195
      Left            =   2760
      TabIndex        =   4
      Top             =   2700
      Width           =   1215
   End
End
Attribute VB_Name = "frmFreedom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QuitDate As Date
Dim CigsPerDay As Integer
Dim PricePer25 As Double
Dim Flash As Integer
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdConfig_Click()
    On Error Resume Next
    Me.Height = 3795
    cmdConfig.Enabled = False
    cboStartDate.Text = Format(QuitDate, "mmmm dd yyyy")
    txtNum = CigsPerDay
    txtPrice = PricePer25
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    If Val(txtNum) = 0 Then MsgBox "Please enter the number of cigarettes smoked per day ": txtNum.SetFocus: Exit Sub
    If Val(txtPrice) = 0 Then MsgBox "Please enter the Price for a pack of 25 in dollars and cents (00.00)": txtPrice.SetFocus: Exit Sub
    
    Me.Height = 2655
    cmdConfig.Enabled = True
    CigsPerDay = CInt(txtNum)
    QuitDate = cboStartDate.Text
    PricePer25 = CDbl(txtPrice)
    SaveSetting App.Title, "Settings", "CigsPerDay", CigsPerDay
    SaveSetting App.Title, "Settings", "QuitDate", QuitDate
    SaveSetting App.Title, "Settings", "PricePer25", PricePer25
    GetDetails
    cmdClose.SetFocus
End Sub

Private Sub Form_Load()
    On Error Resume Next
    For i = 60 To 0 Step -1     'range of approx 2 months back from today
        cboStartDate.AddItem Format(Now - i, "mmmm dd yyyy")
    Next
    cboStartDate.ListIndex = 0
    
    CigsPerDay = GetSetting(App.Title, "Settings", "CigsPerDay", 0)
    If CigsPerDay = 0 Then
        Me.Height = 3795
        cmdConfig.Enabled = False
        lblInfo = "First Time Use: Please enter your Quit Date and smoking quota."
        imgDay(0).Picture = img(0).Picture
        Timer1.Enabled = False
    Else
        QuitDate = GetSetting(App.Title, "Settings", "QuitDate", Now)
        PricePer25 = GetSetting(App.Title, "Settings", "PricePer25", 0)
        Me.Height = 2655
        GetDetails
    End If
End Sub

Private Sub Timer1_Timer()
    imgCongrats.Visible = Not imgCongrats.Visible
    Flash = Flash + 1
    'stop flashing before it gets irritating..
    If Flash >= 6 Then
        imgCongrats.Visible = True
        Timer1.Enabled = False
        lblMe.Visible = True
    End If
End Sub

'didn't bother trapping clipboard pastes..
Private Sub txtNum_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57       'allow digits
        Case 8              'allow backspace
        Case Else
            KeyAscii = 0
    End Select
End Sub

'didn't bother trapping clipboard pastes..
Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57       'allow digits
        Case 8              'allow backspace
        Case 46
            If InStr(txtPrice, ".") > 0 Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Public Sub GetDetails()
    Dim nDays As Integer
    nDays = DateDiff("d", QuitDate, Now)
    n1 = nDays * CigsPerDay
    n2 = n1 / 25
    n3 = n2 * PricePer25
    lblInfo = "Smoke-free days = " & nDays & ", Cigarettes NOT smoked = " & n1 & " (" & Format(n2, "0.00") & " packs), Money saved so far = $" & Format(n3, "00.00")
    'build up number display using images (plain numbers are too dull)..
    Select Case nDays
        Case Is < 10
            imgDay(0).Picture = img(nDays).Picture
        Case Is < 100
            imgDay(0).Picture = img(Int(Mid(Trim(Str(nDays + 1)), 1, 1))).Picture
            imgDay(1).Picture = img(Int(Mid(Trim(Str(nDays + 1)), 2, 1))).Picture
        Case Is < 1000
            imgDay(0).Picture = img(Int(Mid(Trim(Str(nDays + 1)), 1, 1))).Picture
            imgDay(1).Picture = img(Int(Mid(Trim(Str(nDays + 1)), 2, 1))).Picture
            imgDay(2).Picture = img(Int(Mid(Trim(Str(nDays + 1)), 3, 1))).Picture
    End Select
    
    Flash = 0
    Timer1.Enabled = True
End Sub
