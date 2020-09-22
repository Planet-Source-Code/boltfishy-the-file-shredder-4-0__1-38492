VERSION 5.00
Begin VB.Form frmChars 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options :: Characters"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   2040
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2775
      Begin VB.TextBox txtChars 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   0
         Text            =   "frmChars.frx":0000
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "TFS uses the characters below to overwrite a file randomly."
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click() 'ok
    MyChars = txtChars
    Unload Me
End Sub


Private Sub Form_Load()
    txtChars.Text = MyChars
End Sub

