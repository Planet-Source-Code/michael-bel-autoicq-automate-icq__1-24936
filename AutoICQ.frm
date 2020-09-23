VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AutoICQ 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "AutoICQ Beta"
   ClientHeight    =   8400
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12165
   Icon            =   "AutoICQ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   23
      Top             =   8130
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:34"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Timers:Active"
            TextSave        =   "Timers:Active"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "My Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   11175
      Begin VB.CommandButton Command6 
         Caption         =   "Load List..."
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save List...."
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Remove All"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Remove Selected"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   2055
      End
      Begin VB.ComboBox Sbox 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Text            =   "0 > Online / Connect"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox tim 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Schedule"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   3975
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   240
         Top             =   1680
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   4320
         TabIndex        =   8
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Status to set"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Time"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Set Status to:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "When:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Programmed Events:"
         Height          =   255
         Left            =   4320
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10080
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoICQ.frx":0442
            Key             =   "time"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoICQ.frx":059E
            Key             =   "status"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoICQ.frx":06FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Messages"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   11175
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   240
         Top             =   1680
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Schedule"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   3975
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   -240
         Top             =   3960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AutoICQ.frx":2EAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AutoICQ.frx":378A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1935
         Left            =   4320
         TabIndex        =   19
         Top             =   960
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3413
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
         ColHdrIcons     =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "UIN to send to:"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Message String"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.TextBox Strtxt 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Text            =   "Do you like scary users, Cindy?"
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox msgtim 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox UINtxt 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Text            =   "86372684"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Programmed Events:"
         Height          =   255
         Left            =   4320
         TabIndex        =   21
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "String to Send:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "When:"
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "UIN to send: "
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   11520
      Picture         =   "AutoICQ.frx":5F3E
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "AutoICQ Beta  by Michael Belenky 2001(c)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7920
      TabIndex        =   22
      Top             =   240
      Width           =   3615
   End
   Begin VB.Menu icq 
      Caption         =   "AutoIcq"
      Begin VB.Menu About 
         Caption         =   "About..."
      End
      Begin VB.Menu quit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Timers 
      Caption         =   "Timers"
      Begin VB.Menu intervals 
         Caption         =   "Intervals..."
      End
      Begin VB.Menu Timers_DI 
         Caption         =   "Disabled"
      End
      Begin VB.Menu Timers_EN 
         Caption         =   "Enabled"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "AutoICQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public msgFileStr As String
Public Timer_Intervals As Integer



Private Sub About_Click()
aboutfrm.Show
End Sub

Private Sub Command1_Click()

ListView1.ListItems.Add , , Sbox.Text, 3, 3
ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , tim.Text, , ""

End Sub

Private Sub Command2_Click()
msgFileStr = InputBox("Enter the full path of the file to be read from, note that only the first 450 chars will be read and sent!", "AutoICQ 0.60", "C:\bootlog.txt")
'Strtxt.Text = "(File) " & msgFileStr
End Sub

Private Sub Command3_Click()
' remove all items from listview1
ListView1.ListItems.Remove (ListView1.SelectedItem.Index)

End Sub

Private Sub Command4_Click()
ListView1.ListItems.Clear

End Sub

Private Sub Command5_Click()
FileName = InputBox("Enter the file name for this list of events:", "AutoICQ Beta", "events.txt")

If FileName = "" Then Exit Sub ' in case user hits Cancel

Open App.Path & "\" & FileName For Append As #1
 
 For i = 1 To ListView1.ListItems.Count
  
   Print #1, ListView1.ListItems(i).Text
   Print #1, ListView1.ListItems(i).ListSubItems(1).Text
   
 Next i

Close #1

End Sub

Private Sub Command6_Click()
Dim b_status, b_time As String 'buffers

FileName = InputBox("Enter the events file to load:", "AutoICQ Beta", "events.txt")
If FileName = "" Then Exit Sub
ListView1.ListItems.Clear

Open App.Path & "\" & FileName For Input As #1

'On Error Resume Next
 While Not EOF(1)
  
  Line Input #1, b_status
  Line Input #1, b_time
 'MsgBox b_status
 'MsgBox b_time
 ListView1.ListItems.Add , , b_status, 2, 2
 ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , b_time, , ""
 
 Wend
 
  

Close #1


End Sub

Private Sub Command7_Click()

ListView2.ListItems.Add , , UINtxt.Text, 1, 1
ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , msgtim.Text, 2, ""
ListView2.ListItems(ListView2.ListItems.Count).ListSubItems.Add , , Strtxt.Text, 2, ""

End Sub

Private Sub Command8_Click()
ListView3.ListItems.Add , , Uintxt2.Text, 1, 1
ListView3.ListItems(ListView3.ListItems.Count).ListSubItems.Add , , timtxt2.Text, 1, ""
ListView3.ListItems(ListView3.ListItems.Count).ListSubItems.Add , , msgFileStr, 1, ""
End Sub

Private Sub Form_Load()

Init_ICQ_API 'Sets license key to be used with the ICQ API (see sub)


Sbox.AddItem "0 >  Online / Connect"
Sbox.AddItem "1 >  Free For Chat"
Sbox.AddItem "2 >  AWAY"
Sbox.AddItem "3 >  N/A"
Sbox.AddItem "4 >  Occupied (urgeny msgs)"
Sbox.AddItem "5 >  DND (do not disturb)"
Sbox.AddItem "6 >  Invisible (privacy)"
Sbox.AddItem "7 >  Offline / Disconnect"


tim.Text = Time
msgtim.Text = Time
'timtxt2.Text = Time


'Global Const BICQAPI_USER_STATE_ONLINE = 0
'Global Const BICQAPI_USER_STATE_CHAT = 1
'Global Const BICQAPI_USER_STATE_AWAY = 2
'Global Const BICQAPI_USER_STATE_NA = 3
'Global Const BICQAPI_USER_STATE_OCCUPIED = 4
'Global Const BICQAPI_USER_STATE_DND = 5
'Global Const BICQAPI_USER_STATE_INVISIBLE = 6
'Global Const BICQAPI_USER_STATE_OFFLINE = 7


End Sub

Public Sub Init_ICQ_API()
  'you should obtain your own license key from Mirabilis --> web.icq.com\api
  sName = "Visual Basic"
  sPassword = "aaaaaaaa"
  sLicense = "E94AD7C14D1DBAE8"
   
  Rtn = SetLicenseKey(sName, sPassword, sLicense)
 
 Debug.Print "SetLK Returned: " & Rtn
  
 ' Rtn = GetVersion(iCQversion) 'Used sometimes to determine the Icq api ver.
 




End Sub

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub intervals_Click()

Timer_Intervals = InputBox("Enter interval for timers: (secs.)", "AutoICQ Beta", "1")

If Str(Timer_Intervals) = "" Then GoTo ex
On Error GoTo ex

Timer1.Interval = Timer_Intervals * 1000
Timer1.Interval = Timer_Intervals * 1000


ex:

End Sub

Private Sub quit_Click()
Unload Me
End
End Sub

Private Sub Timer1_Timer()
'On Error Resume Next
 For i = 1 To ListView1.ListItems.Count
   
   If ListView1.ListItems(i).SubItems(1) = Time Then
     SetOwnerState (Int(Val(Left(ListView1.ListItems(i).Text, 1))))
    End If
    
 Next i





End Sub

Private Sub Timer2_Timer()
'Hello, Cindy!

For i = 1 To ListView2.ListItems.Count

 If ListView2.ListItems(i).SubItems(1) = Time Then
  
  ' If Left(ListView2.ListItems(i).SubItems(2), 1) = "(" Then
  '  'MsgBox msgFileStr
  '
  '  'MsgBox ("TODO: CODE TO READ FROM THE FILE, msgFileStr")
  '   sent = SendFile(Int(Val(ListView2.ListItems(i))), msgFileStr)
  '
   
  '  Else
  '
    'the string to send is in ListView2.ListItems(i).ListSubItems(2)
     sent = SendMessage(Int(Val(ListView2.ListItems(i))), "_")
    'after the send dialog is called I'm using sendkeys to put the
    'string to be sent and to confirm, the string should be sent without
    'using sentkeys in sendmessage call, but it does not seem to work
    'so I do it this way, if u find out why, please tell me
     
     SendKeys ListView2.ListItems(i).ListSubItems(2)
     SendKeys "%(s)", True '!! Tell the Icq client to confirm the send
     
     Debug.Print "SendMessage rtn: " & sent
   
   End If
'End If

Next i

End Sub

Private Sub Timer3_Timer()
For i = 1 To ListView2.ListItems.Count
 
 If ListView2.ListItems(i).SubItems(1) = Time Then
  
  gih = SendFile(Val(ListView3.ListItems(i)), ListView3.ListItems(i).ListSubItems(2).Text)
  
 End If
Next i
End Sub

Private Sub Timers_DI_Click()
MsgBox "Warning! you are about to disable timers, scheduled events will not be executed", vbExclamation
Timer1.Enabled = False
Timer2.Enabled = False
sb1.Panels(2).Text = "Timers:OFF"
Timers_DI.Checked = True
Timers_EN.Checked = False
End Sub

Private Sub Timers_EN_Click()
Timer1.Enabled = True
Timer2.Enabled = True
sb1.Panels(2).Text = "Timers:Active"
Timers_DI.Checked = False
Timers_EN.Checked = True
End Sub
