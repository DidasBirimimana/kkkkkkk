VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Register 
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1800
      Top             =   5400
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Documents\vb programming\BIT3\Shop.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Documents\vb programming\BIT3\Shop.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "User"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   735
      Left            =   4200
      TabIndex        =   12
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   735
      Left            =   1440
      TabIndex        =   11
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox passwd 
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   3600
      Width           =   2895
   End
   Begin VB.TextBox uname 
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox txtaddress 
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox lname 
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox fname 
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Password"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Username"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Address"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Last name"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "First name"
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "REGISTER FORM"
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ' Declare variables for the input fields
    Dim Firstname As String, Lastname As String, Address As String, username As String, password As String

    ' Retrieve input values from textboxes
    Firstname = fname.Text
    Lastname = lname.Text
    Address = txtaddress.Text
    username = uname.Text
    password = passwd.Text

    ' Validate required fields
    If fname.Text = "" Or lname.Text = "" Or txtaddress.Text = "" Or uname.Text = "" Or passwd.Text = "" Then
        MsgBox "Please fill all fields before registering.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Database connection string
    Dim conn As Object
    Dim cmd As Object
    On Error GoTo ErrorHandler

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Group 2&1\bit\Shop.mdb;Persist Security Info=False"
    conn.Open

    ' Create the SQL command
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "INSERT INTO [User] ([First name], [Last name], Address, Username, [Password]) " & _
                       "VALUES (?, ?, ?, ?, ?)"
        ' Append parameters
        .Parameters.Append .CreateParameter("firstname", 200, 1, 50, Firstname) ' 200 = adVarChar, 1 = adParamInput
        .Parameters.Append .CreateParameter("lastname", 200, 1, 50, Lastname)
        .Parameters.Append .CreateParameter("Adrress", 200, 1, 50, Address)
        .Parameters.Append .CreateParameter("username", 200, 1, 20, username)
        .Parameters.Append .CreateParameter("password", 200, 1, 20, password)
        ' Execute the command
        .Execute
    End With

    MsgBox "User registered successfully!", vbInformation, "Registration Complete"
      Login.Show
      Me.Hide
      
    ' Navigate to the login form
    LoginForm.Show ' Assuming the login form is named LoginForm

    ' Close the registration form if desired
    Me.Hide

    ' Clean up
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during registration: " & Err.Description, vbCritical, "Error"
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    Set cmd = Nothing
    Set conn = Nothing
End Sub

Private Sub CancelBtn_Click()
    ' Clear all input fields
    ClearInputFields

    ' Optionally, display a message confirming the cancellation
    MsgBox "Input fields have been cleared.", vbInformation, "Canceled"
End Sub

' Subroutine to clear all input fields
Private Sub ClearInputFields()
    fname.Text = ""
    lname.Text = ""
    txtaddress.Text = "" ' Corrected email field
    uname.Text = ""
    passwd.Text = ""
End Sub


