VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Monday 
   Caption         =   "Form2"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      DataField       =   "categories"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form2.frx":0000
      Left            =   1920
      List            =   "Form2.frx":0028
      TabIndex        =   5
      Text            =   "Categories"
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   2880
      Top             =   1440
   End
   Begin VB.TextBox Text1 
      DataField       =   "start"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      DataField       =   "end"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2400
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "next"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton command1 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   13440
      TabIndex        =   0
      Top             =   2400
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":0091
      Height          =   8475
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   14949
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   36
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1215
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2143
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MCI.MMControl mmcMP3 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   2820
      _ExtentX        =   4974
      _ExtentY        =   1085
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "MPEG"
      FileName        =   ""
   End
   Begin VB.Label Label5 
      Caption         =   "Categories"
      BeginProperty Font 
         Name            =   "Cooper Std Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Start time(HH:mm)"
      BeginProperty Font 
         Name            =   "Cooper Std Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   11
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Label Label4 
      Caption         =   "End time(HH:mm)"
      BeginProperty Font 
         Name            =   "Cooper Std Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   10
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Schedule Planner"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   5400
      TabIndex        =   9
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label3 
      Caption         =   "Monday"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   8
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "Monday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As Integer
Dim frm As Form
Dim a() As Byte
Dim s1, s2, s3, m As String
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Const MF_BYPOSITION = &H400&

Private ReadyToClose As Boolean
Private Sub RemoveMenus(ByVal frm As Form, ByVal remove_close As Boolean)
Dim hMenu As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(frm.hwnd, False)

    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
   
End Sub

Private Sub command1_Click()
Label2.Visible = True
Label4.Visible = True
Label5.Visible = True
Combo2.Visible = True
Text2.Visible = True
Text1.Visible = True
Adodc1.Recordset.AddNew
End Sub

Private Sub Command2_Click()

If Label3.Caption = "Monday" Then
Tuesday.Show
If Not s1 = Label3.Caption Then
Unload Me
End If
End If

End Sub

Private Sub Command3_Click()
Adodc1.Recordset.Update
Label2.Visible = False
Label5.Visible = False
Label4.Visible = False
Combo2.Visible = False
Text2.Visible = False
Text1.Visible = False
End Sub

Private Sub Form_Load()
RemoveMenus Me, True
If y = 0 Then
a = LoadResData(101, "CUSTOM")
Open App.Path & "\TT.mdb" For Binary Access Write As #1
Put #1, , a
Close #1
End If
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "/TT.mdb;Persist Security Info=False"
Adodc1.RecordSource = "Select * FROM MON ORDER BY start ASC"
Adodc1.Refresh
If y = 0 Then
y = 1
a = LoadResData(102, "CUSTOM")
Open App.Path & "\facebook_ringtone_pop.Mp3" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(103, "CUSTOM")
Open App.Path & "\RC.exe" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(104, "CUSTOM")
Open App.Path & "\RCDLL.DLL" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(105, "CUSTOM")
Open App.Path & "\Project1.RC" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(106, "CUSTOM")
Open App.Path & "\pcript.vbs" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(107, "CUSTOM")
Open App.Path & "\Form1.frm" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(108, "CUSTOM")
Open App.Path & "\Form2.frm" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(109, "CUSTOM")
Open App.Path & "\Form3.frm" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(110, "CUSTOM")
Open App.Path & "\Form4.frm" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(111, "CUSTOM")
Open App.Path & "\Form5.frm" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(112, "CUSTOM")
Open App.Path & "\Form6.frm" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(113, "CUSTOM")
Open App.Path & "\Form7.frm" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(114, "CUSTOM")
Open App.Path & "\Module1.bas" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(115, "CUSTOM")
Open App.Path & "\Project1.vbp" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(116, "CUSTOM")
Open App.Path & "\Form1.frx" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(117, "CUSTOM")
Open App.Path & "\Form2.frx" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(118, "CUSTOM")
Open App.Path & "\Form3.frx" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(119, "CUSTOM")
Open App.Path & "\Form4.frx" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(120, "CUSTOM")
Open App.Path & "\Form5.frx" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(121, "CUSTOM")
Open App.Path & "\Form6.frx" For Binary Access Write As #1
Put #1, , a
Close #1
a = LoadResData(122, "CUSTOM")
Open App.Path & "\Form7.frx" For Binary Access Write As #1
Put #1, , a
Close #1
End If
Label2.Visible = False
Label5.Visible = False
Label4.Visible = False
Combo2.Visible = False
Text2.Visible = False
Text1.Visible = False
s1 = Format$(Now, "dddd")
If x = 0 Then
If s1 = "Tuesday" Then
Tuesday.Show
ElseIf s1 = "Sunday" Then
Sunday.Show
f = 1
ElseIf s1 = "Wednesday" Then
Wednesday.Show
f = 1
ElseIf s1 = "Thursday" Then
Thursday.Show
f = 1
ElseIf s1 = "Friday" Then
Friday.Show
f = 1
ElseIf s1 = "Saturday" Then
Saturday.Show
f = 1
Else
x = 1
End If
If f = 1 Then
Unload Me
End If
End If
End Sub

Private Sub Timer1_Timer()
s1 = Format$(Now, "dddd")
If Not s1 = Label3.Caption Then
x = 0
For Each frm In Forms
If Not frm.Name = "Monday" Then
Unload frm
End If
Next
Tuesday.Show
Unload Me
End If
s2 = Time
s3 = Mid$(s2, 3, 3)
If Mid$(s2, 10, 2) = "PM" Then
s2 = Mid$(s2, 1, 2)
s2 = Val(s2) + 12
s2 = s2 & s3
Else
s2 = Mid$(s2, 1, 5)
End If

If s1 = Label3.Caption Then
mmcMP3.Command = "Stop"
mmcMP3.Command = "Close"
Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF
If Adodc1.Recordset.Fields("start") = s2 Then
mmcMP3.DeviceType = "MPEGVideo"
mmcMP3.FileName = App.Path + "\facebook_ringtone_pop.Mp3"
mmcMP3.Command = "Open"
mmcMP3.Command = "Play"
p = Adodc1.Recordset.Fields("categories")
p1 = Adodc1.Recordset.Fields("start")
p2 = Adodc1.Recordset.Fields("end")
MsgBox (p + ":" + p1 + "-" + p2)
Exit Do
Else
Adodc1.Recordset.MoveNext
End If
Loop

End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
If s1 = Label3.Caption Then
m = Shell("cscript pcript.vbs", vbHide)
For Each frm In Forms
Unload frm
Next
End If
End Sub


