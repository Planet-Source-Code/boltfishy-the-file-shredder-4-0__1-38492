VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The File Shredder v4.00"
   ClientHeight    =   3915
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   285
      Left            =   2640
      TabIndex        =   8
      Top             =   3600
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   3600
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Path"
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6090
      Begin VB.ListBox lstFiles 
         Appearance      =   0  'Flat
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
         Height          =   1950
         Left            =   120
         OLEDropMode     =   1  'Manual
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   360
         Width           =   5850
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   6090
      Begin VB.CommandButton cmdClearItem 
         Caption         =   "Clear Item"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear List"
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5040
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options"
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdDeleteAll 
         Caption         =   "Delete All"
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmMain.frx":0442
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSelectFile 
         Caption         =   "&Select File"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearItem 
         Caption         =   "Clear &Item"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuClearList 
         Caption         =   "&Clear List"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteAll 
         Caption         =   "&Delete All"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuCorruption 
         Caption         =   "Corruption Settings"
      End
      Begin VB.Menu mnuChars 
         Caption         =   "Characters"
      End
      Begin VB.Menu mnuS6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMode 
         Caption         =   "Select Mode"
         Begin VB.Menu mnuQuick 
            Caption         =   "Quick"
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuNormal 
            Caption         =   "Normal"
            Checked         =   -1  'True
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnuParanoid 
            Caption         =   "Paranoid"
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuS5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCustom 
            Caption         =   "Edit - Custom"
            Shortcut        =   {F9}
         End
      End
      Begin VB.Menu mnuS4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddContext 
         Caption         =   "Add Context Menu"
      End
      Begin VB.Menu mnuRemoveContext 
         Caption         =   "Remove Context Menu"
      End
   End
   Begin VB.Menu mnuSearch2 
      Caption         =   "Search"
      Begin VB.Menu mnuSearch 
         Caption         =   "&Find Files"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "SYSTEM TRAY"
      Visible         =   0   'False
      Begin VB.Menu sysRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu sysS1 
         Caption         =   "-"
      End
      Begin VB.Menu sysExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------

Private Sub Browse() 'browse for files

Dim File1 As String
CD1.ShowOpen
File1 = FreeFile
    
    If CD1.FileName <> "" Then 'if file name is true
        File1 = CD1.FileName 'return file path
            ElseIf CD1.FileName = "" Then
            File1 = ""
        Exit Sub
    End If
    
    If FileExists(File1) = True Then 'if file exists
        lstFiles.AddItem (File1) 'add it to list of files
            ElseIf FileExists(File1) = False Then
                File1 = ""
            CD1.FileName = ""
        Exit Sub
    End If
    
File1 = ""
CD1.FileName = "" 'replace this so that next time we get a clean space

End Sub

Private Sub cmdClear_Click() 'clear list

    With lstFiles
        .Clear
        .Refresh
    End With

End Sub

Private Sub cmdClearItem_Click()

On Error Resume Next

lstFiles.RemoveItem lstFiles.ListIndex 'remove the select item from list
    
End Sub

Private Sub cmdDeleteAll_Click()

On Error GoTo ErrSub

Dim i As Integer 'counter to go from 1 to no of files
Dim b As Integer 'no of files
Dim File2Del As String 'file to delete
Dim msg As String 'message box

msg = "WARNING: Files cannot be recovered once deleted!"
msg = msg & vbCrLf & "Are you sure?" 'check if sure

    If MsgBox(msg, vbExclamation + vbYesNo, "Sure?") = vbNo Then
        Exit Sub
            Else 'if answer is yes then go ahead
                b = lstFiles.ListCount

    For i = 0 To b - 1 Step 1 'i = 1 to number of files
        frmMain.Enabled = False
            FileTemp = lstFiles.List(i) 'set global filetemp to the file to be deleted
        ShredFile (lstFiles.List(i)) 'kill item - file - on list, get the file from i
    Next i

    If i = b Then SB1.Panels(1) = "Deleted " & b & " files!" 'when finished i.e. when i has reached b (the no. of files)
        frmMain.Enabled = True 're-enable the main form
            lstFiles.Clear 'and clear the list

ErrSub:
Exit Sub

    End If
    
End Sub

Private Sub cmdExit_Click()
    Unload Me: End
End Sub

Private Sub Form_Load()
   
NumberOfTimes = 8 'default value of overwriting
Binary = True 'do binary? sure!
MyChars = "10" 'chars = 0 and 1 i.e. binary
Method = "ShredFile"
Rename = True
Setting = "Normal"

Me.Show
Me.Refresh

'user might have changed app's path so create a new context menu just to be sure

If CreateContextMenu = True Then
    ElseIf CreateContextMenu = False Then
    MsgBox ("There was an error writing to the registry." & vbCrLf & "Sorry, but we can't add a 'ShredFile' option to the context menus."), vbCritical + vbOKOnly, "Error"
End If

'add to listbox

If Command <> "" Then
    lstFiles.AddItem (Command)
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


Private Sub lstFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    'files are dragged and dropped onto the list box

On Error Resume Next

Dim intFiles As Integer
Dim intLenFile As Integer
    
Dim intX As Integer
Dim strFilePath As String

DoEvents
intFiles = Data.Files.Count
    
    For intX = 1 To intFiles '1 to no of files dropped, i.e. keep adding until all the files have been added
        If (GetAttr(Data.Files(intX)) And vbDirectory) = vbDirectory Then 'check for directory
            Exit Sub 'if a dir was dropped then stop
                Else 'but if a file(s) was dropped then
                    intLenFile = Len(Data.Files(intX))
                strFilePath = Left(Data.Files(intX), intLenFile) 'return file path
            lstFiles.AddItem "" & strFilePath 'add the file path to the listbox
        End If
    Next intX

End Sub

'menu and simple links section
'menu bit basically just calls button click events

    Private Sub cmdAbout_Click()
    frmAbout.Show
    End Sub

    Private Sub cmdOptions_Click()
    PopupMenu mnuOptions 'popup the menu command
    End Sub

    Private Sub mnuAbout_Click()
    frmAbout.Show
    End Sub

    Private Sub mnuAddContext_Click()

    If CreateContextMenu = True Then 'success
        MsgBox ("A context menu was added successfully"), vbInformation + vbOKOnly, "Context Menu"
    
    ElseIf CreateContextMenu = False Then 'failure
        MsgBox ("There was an error writing to the registry." & vbCrLf & "Sorry, but we can't add a 'Shred File' option to the context menus."), vbCritical + vbOKOnly, "Error"
    
    End If

    End Sub

    Private Sub mnuChars_Click()
    frmChars.Show
    End Sub

    Private Sub mnuClearItem_Click()
    Call cmdClearItem_Click
    End Sub

    Private Sub mnuClearList_Click()
    Call cmdClear_Click
    End Sub

    Private Sub mnuCorruption_Click()
    frmCustom.Show
    End Sub

    Private Sub mnuDeleteAll_Click()
    Call cmdDeleteAll_Click
    End Sub

    Private Sub mnuExit_Click()
    'Shell_NotifyIcon NIM_DELETE, nid 'remove from tray
    Unload Me: End
    End Sub

    Private Sub mnuOverWrite_Click()
    frmOptions.Show
    End Sub

    Private Sub mnuRemoveContext_Click()
    
    If DeleteContextMenu = True Then
        MsgBox ("The context menu was removed successfully"), vbInformation + vbOKOnly, "Context Menu"
    
    ElseIf DeleteContextMenu = False Then
        MsgBox ("There was an error writing to the registry." & vbCrLf & "Sorry, but we can't remove the 'Shred File' option from the context menus."), vbCritical + vbOKOnly, "Error"
    
    End If
    End Sub

    Private Sub mnuSearch_Click()
    frmSearch.Show
    End Sub

    Private Sub mnuSelectFile_Click()
    Browse
    End Sub

    Private Sub sysExit_click()
    Unload Me: End
    End Sub
    
    Private Sub mnuQuick_Click()
    mnuQuick.Checked = True: mnuNormal.Checked = False: mnuParanoid.Checked = False: mnuCustom.Checked = False
    
    Rename = True
    NumberOfTimes = 3
    Setting = "Quick"
    
    End Sub
    
    Private Sub mnuNormal_Click()
    mnuQuick.Checked = False: mnuNormal.Checked = True: mnuParanoid.Checked = False: mnuCustom.Checked = False
    
    Rename = True
    NumberOfTimes = 8
    Setting = "Normal"

    End Sub
    
    Private Sub mnuParanoid_Click()
    mnuQuick.Checked = False: mnuNormal.Checked = False: mnuParanoid.Checked = True: mnuCustom.Checked = False
    
    Rename = True
    NumberOfTimes = 15
    Setting = "Paranoid"

    End Sub

    Private Sub mnuCustom_Click()
    Setting = "Custom"
    frmCustom.Show
    End Sub
