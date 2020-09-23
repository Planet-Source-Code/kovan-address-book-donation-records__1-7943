VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDonation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Donation Records"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame franames 
      Caption         =   "Names"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   4320
      TabIndex        =   24
      Top             =   0
      Width           =   2775
      Begin VB.ListBox lstnames 
         Height          =   2400
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame fraMenuSys 
      Height          =   735
      Left            =   0
      TabIndex        =   19
      Top             =   5880
      Width           =   7335
      Begin VB.CommandButton cmddonor 
         Caption         =   "DONOR INFO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "REPORT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cdmstart 
         Caption         =   "START MENU"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Date and Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   1440
      TabIndex        =   12
      Top             =   3720
      Width           =   4095
      Begin VB.ListBox lstDonationDate 
         Height          =   1815
         Left            =   840
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.ListBox lstAmount 
         Height          =   1815
         Left            =   2880
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.ListBox lstDonationID 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame framenu 
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   2880
      Width           =   7215
      Begin VB.CommandButton cmdCancel 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "SAVE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "EDIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Donation Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24444929
         CurrentDate     =   36584
      End
      Begin VB.TextBox txtDonorID 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtDonationID 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtDonationAmount 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lbldate 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   7
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label LBLdonationInfo 
         AutoSize        =   -1  'True
         Caption         =   "Donor ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   6
         Top             =   360
         Width           =   780
      End
      Begin VB.Label LBLdonationInfo 
         AutoSize        =   -1  'True
         Caption         =   "Donation ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   5
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label LBLdonationInfo 
         AutoSize        =   -1  'True
         Caption         =   "Donation Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2160
         TabIndex        =   4
         Top             =   1320
         Width           =   1470
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSelection 
      Caption         =   "&Selection"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
      End
   End
End
Attribute VB_Name = "frmDonation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim donarRecords As Recordset
Dim donationRecords As Recordset

Private Sub Calendar1_Click()

End Sub



Private Sub cdmstart_Click()
Me.Hide
frmStart.Show

End Sub

Private Sub cmdAdd_Click()
If cmdAdd.Caption = "ADD" Then
    cmdAdd.Caption = "ADD IT"
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    cmdCancel.Visible = True
    Selector txtDonationAmount
    DTPicker1.Enabled = True
    txtDonationAmount.Locked = False
Else
    msg = MsgBox("Are you sure you want to add this record?", vbYesNo)
    If msg = vbNo Then
        Exit Sub
    End If
    
    Set donorInformation = OpenDatabase(App.Path & "/" & "OurMosque97.mdb")
    SQL = ""
    SQL = "donationAmount"
    Set donationRecords = donorInformation.OpenRecordset(SQL)
    With donationRecords
                .MoveFirst
                .AddNew
                !DonorID = lstNames.ItemData(lstNames.ListIndex)
                !DonationDate = DTPicker1.Value
                !DonationAmount = txtDonationAmount.Text
                .Update
    End With
            donorInformation.Close
    msg = MsgBox("Would you like to add another transaction?", vbYesNo)
    If msg = vbNo Then
        cmdAdd.Caption = "ADD"
        cmdCancel.Visible = False
        cmdDelete.Enabled = True
        cmdEdit.Enabled = True
        txtDonationAmount.Locked = True
        Exit Sub
    End If
    
End If
End Sub

Private Sub cmdCancel_Click()
cmdAdd.Caption = "ADD"
cmdEdit.Enabled = True
cmdCancel.Visible = False
cmdDelete.Enabled = True
txtDonationAmount.Locked = True
End Sub

Private Sub cmdDelete_Click()
msg = MsgBox("Are you sure you want to delete this?", vbYesNo)
If msg = vbNo Then
    Exit Sub
End If

    Set donorInformation = OpenDatabase(App.Path + "\" + "OurMosque97.mdb")

    SQL = ""
    SQL = "DELETE FROM donationAmount WHERE DonorID = " & lstNames.ItemData(lstNames.ListIndex) & " AND " & _
    "DonationID = " & Val(txtDonationID.Text)
    donorInformation.Execute (SQL)
    lstDonationID.RemoveItem (lstDonationID.ListIndex)
    lstAmount.RemoveItem (lstAmount.ListIndex)
    lstDonationDate.RemoveItem (lstDonationDate.ListIndex)
    txtDonationAmount.Text = ""
    txtDonationID.Text = ""
    txtDonorID.Text = ""
    lstAmount.ListIndex = 0
End Sub

Private Sub cmddonor_Click()
Me.Hide
frmDonorInfo.Show
End Sub

Private Sub cmdEdit_Click()
If cmdEdit.Caption = "EDIT" Then
    cmdEdit.Caption = "CANCEL"
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    txtDonationAmount.Locked = False
    DTPicker1.Enabled = True
    Selector txtDonationAmount
    cmdSave.Enabled = True
Else
    cmdEdit.Caption = "EDIT"
    cmdAdd.Enabled = True
    cmdDelete.Enabled = True
    txtDonationAmount.Locked = False
    DTPicker1.Enabled = False
    cmdSave.Enabled = False
    Exit Sub
End If
'Set donorInformation = OpenDatabase(App.Path & "/" & "OurMosque97.mdb")
'SQL = ""
'SQL = "donationAmount"
'Set donationRecords = donorInformation.OpenRecordset(SQL)
'
'With donationRecords
'            .MoveFirst
'            .Edit
'            !DonorID = lstNames.ItemData(lstNames.ListIndex)
'            !DonationDate = txtdate
'            !DonationAmount = txtDonationAmount.Text
'            .Update
'End With
'        donorInformation.Close
'
       
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdReport_Click()
Me.Hide
frmReport.Show

End Sub

Private Sub cmdSave_Click()
Set donorInformation = OpenDatabase(App.Path & "/" & "OurMosque97.mdb")
Set DonorRecords = donorInformation.OpenRecordset("donorInfo")
Do Until DonorRecords.EOF
    If lstNames.Text = DonorRecords!LastName & ", " & DonorRecords!FirstName Then
        lstNames.RemoveItem (lstNames.ListIndex)
        lstNames.AddItem DonorRecords.Fields("LastName") & ", " & DonorRecords.Fields("FirstName")
        
        With DonorRecords
            .Edit
            'statements
            .Update
        End With
    End If
    DonorRecords.MoveNext
Loop
txtDonationAmount.Locked = True
DTPicker1.Enabled = fasle
cmdDelete.Enabled = True
cmdAdd.Enabled = True
cmdEdit.Caption = "EDIT"
cmdSave.Enabled = False
End Sub

Private Sub Form_Load()
'opening the database and table
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

Private Sub lstAmount_Click()
'selects other approperiate fields in list boxes
txtDonationAmount.Text = lstAmount.Text
lstDonationDate.ListIndex = lstAmount.ListIndex
lstDonationID.ListIndex = lstAmount.ListIndex
End Sub

Private Sub lstDonationDate_Click()
'selects other approperiate fields in list boxes
lstAmount.ListIndex = lstDonationDate.ListIndex
lstDonationID.ListIndex = lstDonationDate.ListIndex

'setting the date picker value ot whats in the listbox for date
DTPicker1.Value = lstDonationDate.Text


End Sub

Private Sub lstDonationID_Click()
'selects other approperiate fields in list boxes
txtDonationID.Text = lstDonationID.Text
lstAmount.ListIndex = lstDonationID.ListIndex
lstDonationDate.ListIndex = lstDonationID.ListIndex
End Sub

Private Sub lstNames_Click()
    'clearing list boxes
    lstDonationDate.Clear
    lstAmount.Clear
    lstDonationID.Clear
    
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


Private Sub mnuAdd_Click()
cmdAdd_Click

End Sub

Private Sub mnuCancel_Click()
cmdCancel_Click
End Sub

Private Sub mnuDelete_Click()
cmdDelete_Click
End Sub

Private Sub mnuEdit_Click()
cmdEdit_Click
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuSave_Click()
cmdSave_Click
End Sub
