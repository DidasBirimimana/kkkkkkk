VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Phone 
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureBox1 
      Height          =   4695
      Left            =   5040
      ScaleHeight     =   4635
      ScaleWidth      =   5715
      TabIndex        =   17
      Top             =   360
      Width           =   5775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2640
      Top             =   6720
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
      RecordSource    =   "Phones"
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
   Begin VB.CommandButton Command4 
      Caption         =   "Delete The Record"
      Height          =   615
      Left            =   6840
      TabIndex        =   16
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Quit The Program "
      Height          =   615
      Left            =   4920
      TabIndex        =   15
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Record"
      Height          =   615
      Left            =   3000
      TabIndex        =   14
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New Record"
      Height          =   615
      Left            =   1080
      TabIndex        =   13
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox txtDiscount 
      Height          =   615
      Left            =   3120
      TabIndex        =   12
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox txtUnitPrice 
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox txtQuantity 
      Height          =   735
      Left            =   3000
      TabIndex        =   10
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txtCategory 
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtPhoneNumber 
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   " Discount "
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "UnitPrice"
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Quantity"
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Category"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Name"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "PhoneNumber"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "PHONE FORM"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Phone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ' Declare variables for the input fields
    Dim PhoneNumber As String, Name As String, Category As String
    Dim Quantity As String, UnitPrice   As String, Discount As String

    ' Retrieve input values from textboxes and combo box
    PhoneNumber = txtPhoneNumber.Text
    Name = txtName.Text
    Category = txtCategory.Text
    Quantity = txtQuantity.Text
    UnitPrice = txtUnitPrice.Text
    Discount = txtDiscount.Text
   
    ' Validate required fields
    If PhoneNumber = "" Or Name = "" Or Category = "" Or Quantity = "" Or UnitPrice = "" Or Discount = "" Then
        MsgBox "Please fill all fields before add new the computer.", vbExclamation, "Missing Information"
        Exit Sub
    End If


    ' Database connection and command objects
    Dim conn As Object
    Dim cmd As Object
    On Error GoTo ErrorHandler ' Add error handling

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Documents\vb programming\BIT3\Shop.mdb;Persist Security Info=False;"
    conn.Open



    ' Prepare the SQL command for inserting data
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "INSERT INTO Phones (PhoneNumber,Name, Category, Quantity, UnitPrice, Discount) " & _
                       "VALUES (?, ?, ?, ?, ?, ?)"
        ' Append parameters
        .Parameters.Append .CreateParameter("PhoneNumber", 200, 1, Len(PhoneNumber), PhoneNumber) ' 200 = adVarChar
        .Parameters.Append .CreateParameter("Name", 200, 1, Len(Name), Name)
        .Parameters.Append .CreateParameter("Category", 200, 1, Len(Category), Category)
        .Parameters.Append .CreateParameter("Quantity", 200, 1, Len(Quantity), Quantity)
        .Parameters.Append .CreateParameter(" UnitPrice ", 200, 1, Len(UnitPrice), UnitPrice)
        .Parameters.Append .CreateParameter("Discount", 200, 1, Len(Discount), Discount)
        
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Phone added successfully!", vbInformation, "Add new Phone Complete"

    ' Clear input fields
    ClearInputFields

    ' Clean up
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    ' Handle errors
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    If Not conn Is Nothing Then conn.Close
    Set cmd = Nothing
    Set conn = Nothing
End Sub

Private Sub CancelBtn_Click()
    ' Clear all input fields
    ClearInputFields

    ' Notify user of cancellation
    MsgBox "Input fields have been cleared.", vbInformation, "Canceled"
End Sub

' Subroutine to clear all input fields
Private Sub ClearInputFields()
    txtPhoneNumber.Text = ""
    txtName.Text = ""
    txtCategory.Text = ""
    txtQuantity.Text = ""
    txtUnitPrice .Text = ""
    txtDiscount.Text = ""
   
End Sub

Private Sub Command3_Click()

    ' Ask the user if they want to log out
    Dim response As VbMsgBoxResult
    response = MsgBox("Do you really want to  Quit this program ?", vbYesNo + vbQuestion, "Quit the program Confirmation")

    ' Check the user's response
    If response = vbYes Then
        ' Show the Welcome form if the user clicked Yes
        Dashboard.Show
        ' Optionally, hide or close the current form
        Me.Hide ' or Me.Close if you want to completely close the current form
    Else
        ' If the user clicked No, do nothing or display a message
        MsgBox "You are still logged in.", vbInformation, "Logout Cancelled"
    End If
End Sub


Private Sub Command4_Click()
    ' Ensure a record is selected
    If Trim(txtPhoneNumber.Text) = "" Then
        MsgBox "Please enter or select a record ID to delete.", vbExclamation, "No Record Selected"
        Exit Sub
    End If

    ' Declare objects and variables
    Dim conn As Object
    Dim cmd As Object
    On Error GoTo ErrorHandler

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Documents\vb programming\BIT3\Shop.mdb;Persist Security Info=False"
    conn.Open

    ' Initialize SQL command for deletion
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "DELETE FROM [Phones] WHERE PhoneNumber = ?"
        .Parameters.Append .CreateParameter("PhoneNumber", 3, 1, 10, CLng(txtPhoneNumber.Text)) ' Ensure ID is treated as a Long
        .Execute
    End With

    ' Notify the user of success
    MsgBox "Record deleted successfully!", vbInformation, "Success"

    ' Cleanup and close connection
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    ' Handle any errors
    MsgBox "Error deleting record: " & Err.Description, vbCritical, "Error"
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    Set cmd = Nothing
    Set conn = Nothing
End Sub

Private Sub Command2_Click()
    ' Database connection and command objects
    Dim conn As Object
    Dim rs As Object ' Recordset for fetching data
    On Error GoTo ErrorHandler

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Documents\vb programming\BIT3\Shop.mdb;Persist Security Info=False;"
    conn.Open

    ' Query to retrieve data from Computers table
    Dim sql As String
    sql = "SELECT PhoneNumber, Name, Category, Quantity, UnitPrice, Discount FROM Phones"

    ' Execute the query
    Set rs = conn.Execute(sql)

    ' Display column headers
    Dim header As String
    header = "PhoneNumber" & vbTab & "Name" & vbTab & "Category" & vbTab & "Quantity" & vbTab & "Unit Price" & vbTab & "Discount"
    PictureBox1.Print header
    
    ' Display data
    Dim y As Single
    Do While Not rs.EOF
        ' Format the data as text with tabs for alignment
        Dim line As String
        line = rs("PhoneNumber") & vbTab & rs("Name") & vbTab & rs("Category") & vbTab & rs("Quantity") & vbTab & _
               Format(rs("UnitPrice"), "Currency") & vbTab & Format(rs("Discount") * 100, "0.00") & "%" ' Correct formatting for discount as percentage

        ' Print each row of data
        PictureBox1.Print line
        y = y + 10 ' Move to the next line
        rs.MoveNext
    Loop

    ' Clean up
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    ' Handle errors
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    
    ' Ensure proper cleanup in case of error
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub

