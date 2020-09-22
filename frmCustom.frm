VERSION 5.00
Begin VB.Form frmCustom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options :: Corruption Settings"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2880
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
   ScaleHeight     =   3375
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Custom Deletion Method"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.CheckBox chkBinary 
         Caption         =   "Use Random Data"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Value           =   2  'Grayed
         Width           =   1815
      End
      Begin VB.CheckBox chkRename 
         Caption         =   "Rename files"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   735
      End
      Begin VB.OptionButton optShredFile 
         Caption         =   "Shred File Technology"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.TextBox txtOverWrite 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "1"
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Routine to use"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   120
         X2              =   2520
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label3 
         Caption         =   "Overwrite data how many times?"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   2520
         Y1              =   2520
         Y2              =   2520
      End
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------

Private Sub Command1_Click() 'cancel
Unload Me
End Sub

Private Sub Command2_Click() 'OK

'rename files?
If chkRename.Value = 1 Then
    Rename = True
ElseIf chkRename.Value = 0 Then
    Rename = False
End If

'number of times
On Error GoTo ErrSub

If Val(txtOverWrite.Text) = 0 Then
    NumberOfTimes = 1
Else
    NumberOfTimes = Val(txtOverWrite.Text)
End If

ErrSub:
If Err.Number = 6 Then: MsgBox ("Number is too large!"), vbCritical + vbOKOnly, "Error": Exit Sub
If Err.Number = 0 Then: Resume Next

Setting = "Custom"

With frmMain
    
.mnuQuick.Checked = False: .mnuNormal.Checked = False: .mnuParanoid.Checked = False: .mnuCustom.Checked = True

End With

Unload Me

End Sub

Private Sub Form_Load()

If Rename = True Then 'rename is true
    chkRename.Value = 1 'value of rename checkbox = 1
ElseIf Rename = False Then
    chkRename.Value = 0
End If

txtOverWrite.Text = NumberOfTimes

End Sub
