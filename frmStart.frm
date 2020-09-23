VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Donation Database by: Kovan Abdulla"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   ControlBox      =   0   'False
   FillColor       =   &H000040C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Height          =   975
      Left            =   -240
      TabIndex        =   3
      Top             =   5400
      Width           =   6855
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FF8080&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   4680
      Width           =   6615
      Begin VB.CommandButton cmdReport 
         BackColor       =   &H00FF8080&
         Caption         =   "REPORTS"
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
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdDonationAmount 
         BackColor       =   &H00FF8080&
         Caption         =   "DONATION AMOUNT"
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdDonorInfo 
         BackColor       =   &H00FF8080&
         Caption         =   "DONOR INFORMATION"
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
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Image Image1 
      Height          =   4425
      Left            =   -240
      Picture         =   "frmStart.frx":0000
      Top             =   240
      Width           =   6840
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDonationAmount_Click()
'donation form is showen
frmDonation.Show
Me.Hide

End Sub

Private Sub cmdDonorInfo_Click()
'donor info form is shown
frmDonorInfo.Show
Me.Hide

End Sub

Private Sub cmdExit_Click()
End ' program ends
End Sub



Private Sub cmdReport_Click()
'report form is shown
frmReport.Show
Me.Hide
End Sub

Private Sub Form_Load()
'loads the form but doesnt show them
Load frmDonorInfo
Load frmDonation
End Sub

