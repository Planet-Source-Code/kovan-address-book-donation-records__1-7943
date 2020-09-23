VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Donation Reports"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7515
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMenuSys 
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   7335
      Begin VB.CommandButton cmdExit 
         Caption         =   "EXIT"
         Height          =   375
         Left            =   5640
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cdmstart 
         Caption         =   "START MENU"
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmddonation 
         Caption         =   "DONATIONS"
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmddonor 
         Caption         =   "DONOR INFO"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraDates 
      Caption         =   "Select Date"
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4095
      Begin MSComCtl2.DTPicker DtEnd 
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36584
      End
      Begin MSComCtl2.DTPicker dtStart 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36584
      End
   End
   Begin VB.Frame fratotal 
      Caption         =   "Total Over A Period"
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   4095
      Begin VB.TextBox txttotal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdTotal 
         Caption         =   "FIND TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   480
         TabIndex        =   7
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label lblReport 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Frame frmNames 
      Caption         =   "Names of Donors"
      Height          =   4695
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   3015
      Begin VB.ListBox lstNames 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Date and Amount"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   4095
      Begin VB.ListBox lstDonationID 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox lstAmount 
         Height          =   1815
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.ListBox lstDonationDate 
         Height          =   1815
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSelection 
      Caption         =   "&Selection"
      Begin VB.Menu mnuCalculate 
         Caption         =   "&Calculate"
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim donarRecords As Recordset
Dim donationRecords As Recordset

Private Sub cdmstart_Click()
Me.Hide
frmStart.Show

End Sub

Private Sub cmddonation_Click()
Me.Hide
frmDonation.Show

End Sub

Private Sub cmddonor_Click()
Me.Hide
frmDonorInfo.Show
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdTotal_Click()
'clearing list boxes
    lstDonationDate.Clear
    lstAmount.Clear
    lstDonationID.Visible = False
    
    'if error encountered, skipp this part of code
    On Error Resume Next
    'opening database based on DONORID
    Set donorInformation = OpenDatabase(App.Path & "/" & "OurMosque97.mdb")
    SQL = ""
    SQL = "SELECT FirstName, LastName, DonationDate, DonationAmount " & _
        "FROM donorInfo, donationAmount " & _
        "WHERE donorInfo.DonorID = donationAmount.DonorID AND " & _
        " donationAmount.DonationDate BETWEEN #" & dtStart.Value & "# " & _
        "AND #" & DtEnd.Value & "# " & _
        "AND donationAmount.DonorID=" & lstNames.ItemData(lstNames.ListIndex)
   
    Set DonorRecords = donorInformation.OpenRecordset(SQL)
    Set donationRecords = donorInformation.OpenRecordset(SQL)

    DonorRecords.MoveFirst

'checking all the records and putting appropriate info
'in appropriate controls
With DonorRecords
    Do Until .EOF
        
            strLine = Format(!DonationDate, "Short Date")
            lstDonationDate.AddItem strLine
            strLine = ""
            strLine = Format(!DonationAmount, "currency")
            lstAmount.AddItem strLine
            
        .MoveNext
    Loop
End With
    
'closing database
donorInformation.Close

'selecting first date in the date list box
lstDonationDate.ListIndex = 0
Dim total As Currency
For X = 0 To lstAmount.ListCount - 1
    total = total + lstAmount.List(X)
Next X
txttotal.Text = Format(total, "currency")
End Sub

Private Sub lstAmount_Click()
'selects other approperiate fields in list boxes

'lstDonationDate.ListIndex = lstAmount.ListIndex
'lstDonationID.ListIndex = lstAmount.ListIndex
End Sub

Private Sub lstDonationDate_Click()
'selects other approperiate fields in list boxes
'lstAmount.ListIndex = lstDonationDate.ListIndex
'lstDonationID.ListIndex = lstDonationDate.ListIndex
End Sub
Private Sub lstDonationID_Click()
'selects other approperiate fields in list boxes

lstAmount.ListIndex = lstDonationID.ListIndex
lstDonationDate.ListIndex = lstDonationID.ListIndex
End Sub
Private Sub Form_Load()
Set donorInformation = OpenDatabase(App.Path & "/" & "OurMosque97.mdb")
Set DonorRecords = donorInformation.OpenRecordset("donorInfo")

'opens the first record in a database
DonorRecords.MoveFirst

'adding all the names in the donor info table into a list box
Do Until DonorRecords.EOF
    lstNames.AddItem DonorRecords.Fields("LastName") & ", " & DonorRecords.Fields("FirstName")
    lstNames.ItemData(lstNames.NewIndex) = DonorRecords.Fields("DonorID")
    DonorRecords.MoveNext

Loop
DonorRecords.Close

'selecting the first name in the list box
lstNames.ListIndex = 0

End Sub

Private Sub lstNames_Click()
 'clearing list boxes
    lstDonationDate.Clear
    lstAmount.Clear
    lstDonationID.Clear
    lstDonationID.Visible = True
    txttotal = ""
    'if error encountered, skipp this part of code
    On Error Resume Next
    'opening database based on DONORID
    Set donorInformation = OpenDatabase(App.Path & "/" & "OurMosque97.mdb")
    SQL = ""
    SQL = "SELECT * FROM donationAmount, donorInfo WHERE donationAmount.donorID=donorInfo.donorID"
    Set DonorRecords = donorInformation.OpenRecordset(SQL)
    Set donationRecords = donorInformation.OpenRecordset(SQL)

    DonorRecords.MoveFirst

'checking all the records and putting appropriate info
'in appropriate controls
With DonorRecords
    Do Until .EOF
        If lstNames.Text = !LastName & ", " & !FirstName Then
            strLine = Format(!DonationDate, "Short Date")
            lstDonationDate.AddItem strLine
            strLine = ""
            strLine = Format(!DonationAmount, "currency")
            lstAmount.AddItem strLine
            lstDonationID.AddItem !donationId
            txtDonorID.Text = lstNames.ItemData(lstNames.ListIndex)
            txtDonationID.Text = !donationId
        End If
        .MoveNext
    Loop
End With
    
'closing database
donorInformation.Close

'selecting first date in the date list box
lstDonationDate.ListIndex = 0

End Sub

Private Sub mnuCalculate_Click()
cmdTotal_Click

End Sub

Private Sub mnuExit_Click()
End

End Sub
