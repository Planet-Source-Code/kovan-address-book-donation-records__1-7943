VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Records"
   ClientHeight    =   4365
   ClientLeft      =   2085
   ClientTop       =   3450
   ClientWidth     =   3915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frasave 
      Height          =   735
      Left            =   0
      TabIndex        =   18
      Top             =   3600
      Width           =   3855
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
         Left            =   2040
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "SAVE DATA"
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
         Width           =   1695
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Donor Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   8
         Left            =   240
         TabIndex        =   20
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtInfo 
         DataSource      =   "dbOurMosque"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   2430
         Width           =   1935
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   7
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtInfo 
         DataSource      =   "dbOurMosque"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label DonorInfo 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
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
         Index           =   8
         Left            =   2400
         TabIndex        =   21
         Top             =   3120
         Width           =   540
      End
      Begin VB.Label DonorInfo 
         AutoSize        =   -1  'True
         Caption         =   "Postal Code"
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
         Left            =   2400
         TabIndex        =   17
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label DonorInfo 
         AutoSize        =   -1  'True
         Caption         =   "Province/State"
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
         Left            =   2400
         TabIndex        =   16
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label DonorInfo 
         AutoSize        =   -1  'True
         Caption         =   "Phone Number"
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
         Left            =   2400
         TabIndex        =   15
         Top             =   2400
         Width           =   1260
      End
      Begin VB.Label DonorInfo 
         AutoSize        =   -1  'True
         Caption         =   "Fax Number"
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
         Index           =   3
         Left            =   2400
         TabIndex        =   14
         Top             =   2760
         Width           =   1020
      End
      Begin VB.Label DonorInfo 
         AutoSize        =   -1  'True
         Caption         =   "Last Name"
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
         Index           =   4
         Left            =   2400
         TabIndex        =   13
         Top             =   600
         Width           =   915
      End
      Begin VB.Label DonorInfo 
         AutoSize        =   -1  'True
         Caption         =   "Address"
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
         Index           =   5
         Left            =   2400
         TabIndex        =   12
         Top             =   960
         Width           =   690
      End
      Begin VB.Label DonorInfo 
         AutoSize        =   -1  'True
         Caption         =   "City"
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
         Index           =   6
         Left            =   2400
         TabIndex        =   11
         Top             =   1320
         Width           =   330
      End
      Begin VB.Label DonorInfo 
         AutoSize        =   -1  'True
         Caption         =   "First Name"
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
         Index           =   7
         Left            =   2400
         TabIndex        =   10
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "&Data"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnudivider 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel"
      End
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DontCheck As Boolean

Private Sub cmdCancel_Click()
'if cancel pressed, everything is reset to normal
frmAdd.Hide
frmDonorInfo.Show
End Sub

Private Sub cmdSave_Click()
'unloading the form
Unload frmDonorInfo

'makes sure there is something in all text boxes
'bad style of code but didn't have time to come up with a way to do it better

If txtinfo(0).Text = "" Then
    MsgBox "Enter Value"
    txtinfo(0).SetFocus
    Exit Sub
ElseIf txtinfo(1).Text = "" Then
    MsgBox "Enter Value"
    txtinfo(1).SetFocus
    Exit Sub
ElseIf txtinfo(2).Text = "" Then
    MsgBox "Enter Value"
    txtinfo(2).SetFocus
    Exit Sub
ElseIf txtinfo(3).Text = "" Then
    MsgBox "Enter Value"
    txtinfo(3).SetFocus
    Exit Sub
ElseIf txtinfo(4).Text = "" Then
    MsgBox "Enter Value"
    txtinfo(4).SetFocus
    Exit Sub
ElseIf txtinfo(5).Text = "" Then
    MsgBox "Enter Value"
    txtinfo(5).SetFocus
    Exit Sub
ElseIf txtinfo(6).Text = "" Then
    MsgBox "Enter Value"
    txtinfo(6).SetFocus
    Exit Sub
ElseIf txtinfo(7).Text = "" Then
    MsgBox "Enter Value"
    txtinfo(7).SetFocus
    Exit Sub
ElseIf txtinfo(8).Text = "" Then
    MsgBox "Enter Value"
    txtinfo(8).SetFocus
    Exit Sub
End If

'if the user hits yes, then it will save the information
msg = MsgBox("Are you sure you want to add this record?", vbYesNo)
If msg = vbYes Then

    'opening the table for update/save
    Set donorInformation = OpenDatabase(App.Path + "\" + "OurMosque97.mdb")
    SQL = ""
    SQL = "SELECT * FROM donorInfo"
    
    'adding a new record to the database
    Set DonorRecords = donorInformation.OpenRecordset(SQL)
     With DonorRecords
            .AddNew
            !LastName = Trim(UCase(txtinfo(0).Text))
            !FirstName = Trim(UCase(txtinfo(1).Text))
            !Address = Trim(UCase(txtinfo(2).Text))
            !City = Trim(UCase(txtinfo(3).Text))
            !PostalCode = Trim(UCase(txtinfo(4).Text))
            !Province = Trim(UCase(txtinfo(5).Text))
            !PhoneNumber = Trim(txtinfo(6).Text)
            !FaxNumber = Trim(UCase(txtinfo(7).Text))
            !Email = Trim(UCase(txtinfo(8).Text))
            .Update
     End With
        donorInformation.Close
        
        'updating the listbox
        frmDonorInfo.lstNames.AddItem Trim(UCase(txtinfo(1))) & ", " & Trim(UCase(txtinfo(0)))
Else
    'if they dont want to save everything resets to normal
    frmAdd.Hide
    frmDonorInfo.Show
    Exit Sub
End If
'clears all the textboxes on a form
deleteTextBoxes frmAdd

'if they dont want to add a new record, reset everything
msg = MsgBox("Do you want to add another record?", vbYesNo)
If msg = vbNo Then
    Unload Me
    frmDonorInfo.Show
End If
End Sub

Private Sub mnuCancel_Click()
cmdCancel
End Sub

Private Sub mnuDonation_Click()
'menu click events happening
Unload Me
frmDonation.Show
End Sub

Private Sub mnuDonor_Click()
Unload Me
frmDonorInfo.Show

End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuOption_Click()

End Sub

Private Sub mnuSave_Click()
cmdSave_Click

End Sub

Private Sub txtinfo_LostFocus(Index As Integer)
   
   'makes sure there is a value in all of them
   If txtinfo(Index) = "" And DontCheck = False Then
      MsgBox "Please enter a value"
      DontCheck = True
      txtinfo(Index).SetFocus
   Else
      DontCheck = False
   End If
   
End Sub
