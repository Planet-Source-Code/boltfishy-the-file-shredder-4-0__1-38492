VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2895
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   3585
   ClipControls    =   0   'False
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1998.18
   ScaleMode       =   0  'User
   ScaleWidth      =   3366.5
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1200
      TabIndex        =   0
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Label Label6 
      Caption         =   "boltfish@eml.cc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Please send all feedback to:"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "http://www26.brinkster.com/boltfish/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Visit our website at:"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright (C) 2002 Mischa Balen."
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "The File Shredder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   676.117
      X2              =   3267.9
      Y1              =   579.783
      Y2              =   579.783
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   676.117
      X2              =   3267.9
      Y1              =   579.783
      Y2              =   579.783
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------
    
    Private Sub cmdOK_Click() 'OK
    Unload Me 'close
    End Sub

    Private Sub Form_Load()
    Me.Caption = "About v" & App.Major & "." & App.Minor
    End Sub
