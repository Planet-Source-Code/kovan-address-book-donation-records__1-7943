VERSION 5.00
Begin VB.Form frmDonorInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Our Mosque Donation Database by: Kovan Abdulla"
   ClientHeight    =   7290
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMenuSys 
      Height          =   735
      Left            =   0
      TabIndex        =   34
      Top             =   4200
      Width           =   7335
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmddonation 
         Caption         =   "DONATION"
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
         TabIndex        =   35
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraList 
      Caption         =   "List of All Donors"
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
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   3015
      Begin VB.ListBox lstNames 
         Height          =   3180
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame frafind 
      Height          =   2175
      Index           =   0
      Left            =   1080
      TabIndex        =   27
      Top             =   5040
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame frafind 
         Height          =   735
         Index           =   1
         Left            =   3240
         TabIndex        =   30
         Top             =   840
         Width           =   1455
         Begin VB.CommandButton cmdFind 
            Caption         =   "FIND"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.ListBox lstFind 
         Height          =   1425
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtfind 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lbllastname 
         AutoSize        =   -1  'True
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3240
         TabIndex        =   29
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.Frame fraMenu 
      Height          =   735
      Left            =   0
      TabIndex        =   26
      Top             =   3480
      Width           =   7335
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
         Left            =   3120
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "SEARCH"
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
         Left            =   6000
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "NEW"
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
         Left            =   4560
         TabIndex        =   4
         Top             =   240
         Width           =   1335
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
         Left            =   3120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdModify 
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
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "NEXT >"
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
         TabIndex        =   1
         Top             =   240
         Width           =   1335
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
      Left            =   3240
      TabIndex        =   17
      Top             =   0
      Width           =   4095
      Begin VB.TextBox txtinfo 
         Height          =   285
         Index           =   8
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtinfo 
         DataSource      =   "dbOurMosque"
         Height          =   285
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtinfo 
         Height          =   285
         Index           =   7
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtinfo 
         Height          =   285
         Index           =   6
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtinfo 
         Height          =   285
         Index           =   5
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtinfo 
         Height          =   285
         Index           =   4
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtinfo 
         Height          =   285
         Index           =   3
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtinfo 
         Height          =   285
         Index           =   2
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtinfo 
         DataSource      =   "dbOurMosque"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   600
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
         TabIndex        =   32
         Top             =   3120
         Width           =   540
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
         TabIndex        =   25
         Top             =   240
         Width           =   915
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
         TabIndex        =   24
         Top             =   1320
         Width           =   330
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
         TabIndex        =   23
         Top             =   960
         Width           =   690
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
         TabIndex        =   22
         Top             =   600
         Width           =   915
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
         TabIndex        =   21
         Top             =   2760
         Width           =   1020
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
         TabIndex        =   20
         Top             =   2400
         Width           =   1260
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
         TabIndex        =   19
         Top             =   2040
         Width           =   1305
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
         TabIndex        =   18
         Top             =   1680
         Width           =   1035
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
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
      End
   End
End
Attribute VB_Name = "frmDonorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim donorInformation As Database
Dim DonorRecords As Recordset
Dim DontCheck As Boolean
'Dim SQL As String


Private Sub cdmstart_Click()
Me.Hide
frmStart.Show

End Sub

Private Sub cmdCancel_Click()
cmdNext.Enabled = True
cmdDelete.Enabled = True
cmdNew.Enabled = True
cmdSearch.Enabled = True
lockTextboxes frmDonorInfo
cmdCancel.Visible = False
cmdModify.Caption = "EDIT"

End Sub

Private Sub cmdDelete_Click()
    'if error find, it will delte the record
    On Error GoTo cantDelete

    Set donorInformation = OpenDatabase(App.Path + "\" + "OurMosque97.mdb")
    SQL = ""
    SQL = "SELECT * FROM donorInfo"
    Set DonorRecords = donorInformation.OpenRecordset(SQL)
        Do Until DonorRecords.EOF
            If lstNames.Text = DonorRecords!LastName & ", " & DonorRecords!FirstName Then
                DonorRecords.Delete
                lstNames.RemoveItem (lstNames.ListIndex)
            End If
            DonorRecords.MoveNext
        Loop
    deleteTextBoxes Me
    lstNames.ListIndex = 0
cantDelete:
     MsgBox ("Record cannot be deleted if other tables are related to this record")


    
    
End Sub

Private Sub cmddonation_Click()
Me.Hide
frmDonation.Show
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdFind_Click()
'On Error Resume Next
Set donorInformation = OpenDatabase(App.Path & "/" & "OurMosque97.mdb")
lstFind.Clear
If txtfind.Text = "" Then
    MsgBox "Please enter last name to be searched"
    txtfind.SetFocus
    Exit Sub
End If

   'SQL = ""
   Dim LName As String
   Dim mySQL As String
   LName = Trim(UCase(txtfind))
   mySQL = "SELECT LastName, FirstName " & _
         "FROM donorInfo " & _
         "WHERE Lastname LIKE '*" & LName & "*'"

Set DonorRecords = donorInformation.OpenRecordset(mySQL)

With DonorRecords
    If .EOF Then
            MsgBox "No matching donor found, try again please"

            Selector txtfind
    Else
        
            Do Until .EOF
                lstFind.AddItem .Fields("LastName") & ", " & .Fields("FirstName")
                .MoveNext
            Loop
    End If
End With
End Sub

Private Sub cmdModify_Click()
cmdNext.Enabled = False
cmdDelete.Enabled = False
cmdNew.Enabled = False
cmdSearch.Enabled = False
If cmdModify.Caption = "EDIT" Then
    UnlockTextBoxes frmDonorInfo
    cmdModify.Caption = "MODIFY"
    Selector txtInfo(0)
    cmdCancel.Visible = True
ElseIf cmdModify.Caption = "MODIFY" Then
    msg = ""
    msg = MsgBox("Are you sure you want to modify this record?", vbYesNo)
    If msg = vbYes Then
        Set donorInformation = OpenDatabase(App.Path & "/" & "OurMosque97.mdb")
        Set DonorRecords = donorInformation.OpenRecordset("donorInfo")
                Do Until DonorRecords.EOF
                    If lstNames.Text = DonorRecords!LastName & ", " & DonorRecords!FirstName Then
                        lstNames.RemoveItem (lstNames.ListIndex)
                        lstNames.AddItem txtInfo(1).Text & ", " & txtInfo(0)
                        With DonorRecords
                            .Edit
                            !FirstName = Trim(UCase(txtInfo(0).Text))
                            !LastName = Trim(UCase(txtInfo(1).Text))
                            !Address = Trim(UCase(txtInfo(2).Text))
                            !City = Trim(UCase(txtInfo(3).Text))
                            !Province = Trim(UCase(txtInfo(5).Text))
                            !PostalCode = Trim(UCase(txtInfo(4).Text))
                            !PhoneNumber = Trim(UCase(txtInfo(6).Text))
                            !FaxNumber = Trim(UCase(txtInfo(7).Text))
                            !Email = Trim(UCase(txtInfo(8).Text))
                            .Update
                        End With
                    End If
                    DonorRecords.MoveNext
                Loop
            Dim Control
        For Each Control In Me.Controls
            If TypeOf Control Is TextBox Then Control.Locked = True
        Next Control

        cmdNext.Enabled = True
        cmdDelete.Enabled = True
        cmdNew.Enabled = True
        cmdSearch.Enabled = True
    Else
        Exit Sub
    End If
End If
End Sub


Private Sub cmdNew_Click()
Unload frmDonorInfo

frmAdd.Show vbModal
End Sub

Private Sub cmdNext_Click()
If lstNames.ListIndex = lstNames.ListCount - 1 Then
    lstNames.ListIndex = 0
Else
    lstNames.ListIndex = (lstNames.ListIndex + 1)
End If

End Sub

Private Sub cmdReport_Click()
Me.Hide
frmReport.Show

End Sub

Private Sub cmdsearch_Click()
'if the user wants to search
'displayes the search area
'if the caption of it is "SEARCH"
If cmdSearch.Caption = "SEARCH" Then
    cmdSearch.Caption = "Hide Search"
    frmDonorInfo.Height = 7910
    frmDonorInfo.Width = 7455
    frafind(0).Visible = True
    txtfind.SetFocus
    txtfind.Locked = False
Else
'if the caption of is not "Hide Search"
'you guessed it, it hides the Search area
    cmdSearch.Caption = "SEARCH"
    Me.Height = 5655
    Me.Width = 7455
    
    'mess cleanin up after a search is made
    txtfind.Text = ""
    lstFind.Clear
End If

End Sub


Private Sub Form_Load()
'procedure call
GetRecords

Me.Height = 5655
End Sub

Private Sub lstFind_Click()
    'listbox that lists all the records found
    'when a user does a search
    txtfind.Text = lstFind.Text
    On Error Resume Next
    Set donorInformation = OpenDatabase(App.Path & "/" & "OurMosque97.mdb")
    Set DonorRecords = donorInformation.OpenRecordset("donorInfo")
    DonorRecords.MoveFirst

'going throu the recordset
With DonorRecords
    Do Until .EOF
        If lstFind.Text = !LastName & ", " & !FirstName Then
            txtInfo(0).Text = !FirstName
            txtInfo(1).Text = !LastName
            txtInfo(2).Text = !Address
            txtInfo(3).Text = !City
            txtInfo(4).Text = !Province
            txtInfo(5).Text = !PostalCode
            txtInfo(6).Text = !PhoneNumber
            txtInfo(7).Text = !FaxNumber
            txtInfo(8).Text = !Email
        End If
        .MoveNext
    Loop
End With
End Sub

Private Sub lstNames_Click()
'list box that will hold all the names of the record owners
'on its click event, appropriate info is extracted from database

    On Error Resume Next
    Set donorInformation = OpenDatabase(App.Path & "/" & "OurMosque97.mdb")
    Set DonorRecords = donorInformation.OpenRecordset("donorInfo")
    DonorRecords.MoveFirst
    
With DonorRecords
    Do Until .EOF
        If lstNames.Text = !LastName & ", " & !FirstName Then
            txtInfo(0).Text = !FirstName
            txtInfo(1).Text = !LastName
            txtInfo(2).Text = !Address
            txtInfo(3).Text = !City
            txtInfo(4).Text = !Province
            txtInfo(5).Text = !PostalCode
            txtInfo(6).Text = !PhoneNumber
            txtInfo(7).Text = !FaxNumber
            txtInfo(8).Text = !Email
        End If
        .MoveNext
    Loop
End With
End Sub
Public Sub GetRecords()
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
    
    lstNames.ListIndex = 0
End Sub

Private Sub mnuCancel_Click()
cmdCancel_Click

End Sub

Private Sub mnuEdit_Click()
cmdEdit_Click
End Sub

Private Sub mnuFind_Click()
cmdFind_Click

End Sub

Private Sub mnuNew_Click()
cmdNew_Click
End Sub

Private Sub mnuSeach_Click()
cmdsearch_Click
End Sub

Private Sub mnuSearch_Click()
cmdsearch_Click
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
'if user presses enter, it basically makes it look like
'the user has hit FIND button
If KeyAscii = 13 Then
    cmdFind_Click
    txtfind.SelStart = 0
    txtfind.SelLength = Len(txtfind.Text)
    txtfind.SetFocus
End If
End Sub

Private Sub txtinfo_LostFocus(Index As Integer)
   'making sure they dont skipp any textboxes left empty
   If txtInfo(Index) = "" And DontCheck = False Then
      MsgBox "Please enter a value"
      DontCheck = True
      txtInfo(Index).SetFocus
   Else
      DontCheck = False
    
   End If
End Sub


