VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "Glue Files"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8820
   Icon            =   "Glue Files.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5640
      Top             =   4440
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Remove"
      Height          =   255
      Left            =   7200
      TabIndex        =   9
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Down"
      Height          =   255
      Left            =   8160
      TabIndex        =   8
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Up"
      Height          =   255
      Left            =   8160
      TabIndex        =   7
      Top             =   3720
      Width           =   615
   End
   Begin MSComctlLib.ProgressBar bar 
      Height          =   135
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5160
      Width           =   6855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C00000&
      Caption         =   "Glue all"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00808080&
      TabIndex        =   2
      Top             =   5040
      Width           =   975
   End
   Begin MSComDlg.CommonDialog add 
      Left            =   3960
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   7695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C00000&
      Caption         =   "Add File"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00808080&
      TabIndex        =   0
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Files to glue:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7935
   End
   Begin MSComDlg.CommonDialog save 
      Left            =   6120
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Final file:  bytes"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   4800
      Width           =   3615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Reading:  bytes"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   4440
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Save to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   5160
      Width           =   735
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu action 
      Caption         =   "Action"
      Begin VB.Menu mnuadd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuclearlist 
         Caption         =   "Clear List"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************
'**     Program Created By Misael        **
'**     This program joins Files         **
'**     of any type ideal to join        **
'**     little movies or music           **
'**     or anything else, enjoy.         **
'******************************************
Private Declare Function ShellExecute Lib _
    "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    Dim buffer As String, buffer2 As String
    

Private Sub Command1_Click()
On Error GoTo handleit
add.CancelError = True
add.ShowOpen

List1.AddItem add.FileName
handleit:
Exit Sub
End Sub

Private Sub Command2_Click()
Dim resp As String
resp = cmdMoveUp_Click(List1)
End Sub

Private Sub Command3_Click()
Dim a As String
On Error Resume Next
Timer1.Enabled = True
bar.Max = List1.ListCount
bar.Min = 0
save.CancelError = True
save.DialogTitle = "Add The File Extension"
save.Filter = "All|*.*|MPG|*.mpg|MPEG|*.mpeg|WAV|*.wav|MP3|*.mp3"
save.ShowSave

Text1.Text = save.FileName

DoEvents
For i = 0 To List1.ListCount - 1
List1.ListIndex = i
DoEvents
a = FreeFile

Open List1.Text For Binary Access Read As a
   buffer = Space$(LOF(a))
   Get a, , buffer
   DoEvents
Close a
DoEvents
buffer2 = buffer2 & buffer


bar.Value = i
Next i

Open Text1.Text For Binary Access Write As #1
   Put #1, , buffer2
   DoEvents
Close #1
bar.Value = 0

Timer1.Enabled = False
buffer = ""
buffer2 = ""

MsgBox "Files Glued Succesfully", vbOKOnly + vbInformation, "Status"
End Sub

Private Sub Command4_Click()
Dim resp As String
resp = cmdMoveDown_Click(List1)
End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim listcurrent As String
listcurrent = List1.ListIndex
List1.RemoveItem List1.ListIndex
List1.ListIndex = listcurrent
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_DblClick()
On Error Resume Next
Dim app As Long

If List1.Text = "" Then

Else
app = ShellExecute(Me.hwnd, _
    vbNullString, _
    List1.Text, _
    VBA.vbNormalFocus, _
    "c:\", _
    1)
End If
End Sub

Private Sub mnuabout_Click()
Load credits
credits.Show 1
End Sub

Private Sub mnuadd_Click()
Command1_Click
End Sub

Private Sub mnuclearlist_Click()
List1.Clear
Text1.Text = ""
End Sub
Private Sub mnuclose_Click()
End
End Sub

Public Function cmdMoveUp_Click(lstMove As ListBox) As Integer
On Error Resume Next
 'not by source
    Dim strTemp1 As String   '-- hold the selected index data temporarily for move
    
    Dim iCnt    As Integer   '-- holds the index of the item to be moved
    
    iCnt = lstMove.ListIndex
    
    If iCnt > -1 Then
         
         strTemp1 = lstMove.List(iCnt)
        
        '-- Add the item selected to one position above the current position
        lstMove.AddItem strTemp1, (iCnt - 1)
        
        '-- remove it from the current position. Note the current position has changed because the add has moved everything down by 1
        lstMove.RemoveItem (iCnt + 1)
        
        '-- Reselect the item that was moved.
             lstMove.Selected(iCnt - 1) = True
    
    End If

End Function
Public Function cmdMoveDown_Click(lstMove As ListBox) As Integer
On Error Resume Next
    Dim strTemp1 As String    '-- hold the selected index data temporarily for move
    Dim iCnt    As Integer    '-- holds the index of the item to be moved
        
    '-- Assign the first index
    iCnt = lstMove.ListIndex
    
    If iCnt > -1 Then
         
         strTemp1 = lstMove.List(iCnt)
        
        '-- Add the item selected to below the current position
        lstMove.AddItem strTemp1, (iCnt + 2)
        
        lstMove.RemoveItem (iCnt)
        
        '-- Reselect the item that was moved.
        lstMove.Selected(iCnt + 1) = True
   End If

End Function

Private Sub Timer1_Timer()
Label2.Caption = "Reading: " & Format(Len(buffer), "###,###,###") & " bytes"
Label3.Caption = "Final file: " & Format(Len(buffer2), "###,###,###") & " bytes"
End Sub
