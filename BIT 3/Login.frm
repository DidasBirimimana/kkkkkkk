VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Login 
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1920
      Top             =   4200
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Group 2&1\BIT 3\Shop.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Group 2&1\BIT 3\Shop.mdb;Persist Security Info=False"
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
   Begin VB.TextBox passwrd 
      Height          =   615
      Left            =   3240
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox uname 
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Username"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "LOGIN FORM"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ' Declare variables for the input fields
    Dim username As String, password As String

    ' Retrieve input values from textboxes
    username = uname.Text
    password = passwrd.Text

    ' Validate if both fields are filled
    If uname.Text = "" Or passwrd.Text = "" Then
        MsgBox "Please enter both Username and Password.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Database connection string
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    On Error GoTo ErrorHandler

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Group 2&1\BIT 3\Shop.mdb;Persist Security Info=False"
    conn.Open

    ' Create the SQL command to check for username and password
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "SELECT * FROM [user] WHERE Username = ? AND [Password] = ?"
        ' Append parameters
        .Parameters.Append .CreateParameter("username", 200, 1, 50, username)
        .Parameters.Append .CreateParameter("password", 200, 1, 50, password)
        
        ' Execute the command and get the recordset
        Set rs = .Execute
    End With

    ' Check if any records were returned (i.e., the user exists with the provided credentials)
    If Not rs.EOF Then
        MsgBox "Login successful!", vbInformation, "Login"
        Dashboard.Show
        Me.Hide
        
        ' You can redirect to another form or open the main application window here
        ' For example: OpenMainForm
    Else
        MsgBox "Invalid username or password. Please try again.", vbCritical, "Login Failed"
    End If

    ' Clean up
    rs.Close
    conn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during login: " & Err.Description, vbCritical, "Error"
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
End Sub

Private Sub CancelBtn_Click()
    ' Clear the username and password fields
    uname.Text = ""
    passwrd.Text = ""

    ' Optionally, you can reset the focus to the username field
    uname.SetFocus
End Sub

