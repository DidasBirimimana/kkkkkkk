VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Computer 
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   3360
      Top             =   6720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Group 2&1\BIT 3\BIT 3\Shop.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Group 2&1\BIT 3\BIT 3\Shop.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Computers"
      Caption         =   "Adodc2"
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
   Begin VB.PictureBox PictureBox1 
      Height          =   3975
      Left            =   5400
      ScaleHeight     =   3915
      ScaleWidth      =   7755
      TabIndex        =   18
      Top             =   720
      Width           =   7815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Go to Phone"
      Height          =   375
      Left            =   8280
      TabIndex        =   17
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete The Record "
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Quit The Program "
      Height          =   495
      Left            =   4680
      TabIndex        =   15
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Record"
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New Record"
      Height          =   495
      Left            =   1440
      MaskColor       =   &H000000FF&
      TabIndex        =   13
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txtDiscount 
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox txtUnitPrice 
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox txtQuantity 
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtCategory 
      Height          =   615
      Left            =   3120
      TabIndex        =   9
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtComputerNumber 
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000A&
      Caption         =   "LIST  OF COMPUTER"
      Height          =   375
      Left            =   8040
      TabIndex        =   19
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000A&
      Caption         =   "Discount"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000A&
      Caption         =   "UnitPrice"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000A&
      Caption         =   "Quantity"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000A&
      Caption         =   "Category"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      Caption         =   "Name"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "ComputerNumber"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "COMPUTER FORM"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Computer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ' Declare variables for the input fields
    Dim ComputerNumber As String, Name As String, Category As String
    Dim Quantity As String, UnitPrice   As String, Discount As String

    ' Retrieve input values from textboxes and combo box
    ComputerNumber = txtComputerNumber.Text
    Name = txtName.Text
    Category = txtCategory.Text
    Quantity = txtQuantity.Text
    UnitPrice = txtUnitPrice.Text
    Discount = txtDiscount.Text
   
    ' Validate required fields
    If ComputerNumber = "" Or Name = "" Or Category = "" Or Quantity = "" Or UnitPrice = "" Or Discount = "" Then
        MsgBox "Please fill all fields before add new the computer.", vbExclamation, "Missing Information"
        Exit Sub
    End If


    ' Database connection and command objects
    Dim conn As Object
    Dim cmd As Object
    On Error GoTo ErrorHandler ' Add error handling

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:"
    conn.Open

    ' Prepare the SQL command for inserting data
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "INSERT INTO Computers (ComputerNumber,Name, Category, Quantity, UnitPrice, Discount) " & _
                       "VALUES (?, ?, ?, ?, ?, ?)"
        ' Append parameters
        .Parameters.Append .CreateParameter("ComputerNumber", 200, 1, Len(ComputerNumber), ComputerNumber) ' 200 = adVarChar
        .Parameters.Append .CreateParameter("Name", 200, 1, Len(Name), Name)
        .Parameters.Append .CreateParameter("Category", 200, 1, Len(Category), Category)
        .Parameters.Append .CreateParameter("Quantity", 200, 1, Len(Quantity), Quantity)
        .Parameters.Append .CreateParameter(" UnitPrice ", 200, 1, Len(UnitPrice), UnitPrice)
        .Parameters.Append .CreateParameter("Discount", 200, 1, Len(Discount), Discount)
        
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Computer added successfully!", vbInformation, "Add new computer Complete"

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
    txtComputerNumber.Text = ""
    txtName.Text = ""
    txtCategory.Text = ""
    txtQuantity.Text = ""
    txtUnitPrice .Text = ""
    txtDiscount.Text = ""
   
End Sub

Private Sub ExitBtn_Click()
    ' Close the form properly
    Unload Me
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
    sql = "SELECT ComputerNumber, Name, Category, Quantity, UnitPrice, Discount FROM Computers"

    ' Execute the query
    Set rs = conn.Execute(sql)

    ' Display column headers
    Dim header As String
    header = "Computer Number" & vbTab & "Name" & vbTab & "Category" & vbTab & "Quantity" & vbTab & "Unit Price" & vbTab & "Discount"
    PictureBox1.Print header
    
    ' Display data
    Dim y As Single
    Do While Not rs.EOF
        ' Format the data as text with tabs for alignment
        Dim line As String
        line = rs("ComputerNumber") & vbTab & rs("Name") & vbTab & rs("Category") & vbTab & rs("Quantity") & vbTab & _
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




Private Sub Command4_Click()
    ' Ensure a record is selected
    If Trim(txtComputerNumber.Text) = "" Then
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
        .CommandText = "DELETE FROM [Computers] WHERE ComputerNumber = ?"
        .Parameters.Append .CreateParameter("ComputerNumber", 3, 1, 10, CLng(txtComputerNumber.Text)) ' Ensure ID is treated as a Long
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

Private Sub Command5_Click()
    ' Ask the user if they want to log out
    Dim response As VbMsgBoxResult
    response = MsgBox("Do you really want to go Phone?", vbYesNo + vbQuestion, "Logout Confirmation")

    ' Check the user's response
    If response = vbYes Then
        ' Show the Welcome form if the user clicked Yes
        Phone.Show
        ' Optionally, hide or close the current form
        Me.Hide ' or Me.Close if you want to completely close the current form
    Else
        ' If the user clicked No, do nothing or display a message
        MsgBox "You are still logged in.", vbInformation, "Logout Cancelled"
    End If
End Sub


