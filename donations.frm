VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDonation 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3390
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Donation Amount"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtDonationAmount 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtDonationID 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Format          =   22872065
         CurrentDate     =   36575
      End
      Begin VB.TextBox txtDonorID 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label LBLdonationInfo 
         AutoSize        =   -1  'True
         Caption         =   "Donation Date"
         Height          =   195
         Index           =   3
         Left            =   2160
         TabIndex        =   6
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label LBLdonationInfo 
         AutoSize        =   -1  'True
         Caption         =   "Donation Amount"
         Height          =   195
         Index           =   2
         Left            =   2160
         TabIndex        =   5
         Top             =   1320
         Width           =   1230
      End
      Begin VB.Label LBLdonationInfo 
         AutoSize        =   -1  'True
         Caption         =   "Donation ID"
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.Label LBLdonationInfo 
         AutoSize        =   -1  'True
         Caption         =   "Donor ID"
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmDonation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
