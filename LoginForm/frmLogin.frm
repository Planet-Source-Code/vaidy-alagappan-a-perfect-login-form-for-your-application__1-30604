VERSION 5.00
Begin VB.Form frmLoginUser 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3285
   ClientLeft      =   900
   ClientTop       =   1515
   ClientWidth     =   4800
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   3975.852
   ScaleMode       =   0  'User
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   240
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   1500
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   372
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1000
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Enter"
      Height          =   372
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1000
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   2295
      Left            =   840
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Login-Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      Top             =   1440
      Width           =   900
   End
End
Attribute VB_Name = "frmLoginUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Contacts tropicalwire@hotmail.com
''This program will help you to create a perfect login
''form for your applications.
''database.

''This project, i have used ADO(ActiveX Data Objects
''Library 2.6)
''If you do not have then use atleast 2.1 or 2.6
''This you can change by selecting the Project Menu and
''then References.

''The name of the database is LoginCheck
''The name of the database table is Login

'________________________________________________________


'This is for connecting to the database, i am
'declaring the connection as CON
Dim CON As Connection

Dim i As Integer 'i am declaring 'i' for counting the wrong passwords
Dim rs As New Recordset
Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdEnter_Click()
'First line of code checks for the value of i being less then 3. i.e. it accepts
'maximum of three wrong passwords
If i > 1 Then
    rs.Open "Select * from Login where name='" & txtName.Text & "' and pass='" & txtPass.Text & "'", CON, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
        MsgBox "InValid User & Password", vbCritical, "Failure"
    Else
       MsgBox "Password Correct", vbInformation, "Success!"
       Unload Me
    End If
    rs.Close
    i = i - 1
Else
    MsgBox "InValid User & Password Unloading", vbCritical, "Failure"
    End
End If
End Sub
Private Sub Form_Initialize()
i = 3     'I am initializing the value of 'i' for checking wrong passwords
''Here i am making an instance of the Connection
Set CON = New Connection

''This is nothing but a connection string for the database
''This string varies for any other database like SQL etc
CON.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.Path & "\LoginCheck.mdb"

End Sub

Private Sub Form_Load()
MsgBox "The username & password are 'JOHN'"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End 'I am not unloading i am ending the program
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
'First line of code will check whether the user is using Ctrl+C or Ctrl+X or Ctrl+V
'If he uses so, the second line of code will make the value of keyascii to carry nothing
If KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 24 Then
    KeyAscii = 0
ElseIf KeyAscii = 13 Then 'This is for using the Enter key
    KeyAscii = 0
    SendKeys "{Tab}"
Else
    'Take a look
    'This is for changing the lower case alphabets to upper case
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 22 Or KeyAscii = 3 Or KeyAscii = 24 Then
    KeyAscii = 0
ElseIf KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{Tab}"
    Call cmdEnter_Click
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
