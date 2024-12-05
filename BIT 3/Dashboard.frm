VERSION 5.00
Begin VB.Form Dashboard 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Logout 
      Caption         =   "Logout"
      Height          =   615
      Left            =   5760
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Phone"
      Height          =   615
      Left            =   3600
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Computer"
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "WELCOME TO OUR SHOP"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    ' Ask the user if they want to log out
    Dim response As VbMsgBoxResult
    response = MsgBox("Do you really want to computer?", vbYesNo + vbQuestion, "Logout Confirmation")

    ' Check the user's response
    If response = vbYes Then
        ' Show the Welcome form if the user clicked Yes
        Computer.Show
        ' Optionally, hide or close the current form
        Me.Hide ' or Me.Close if you want to completely close the current form
    Else
        ' If the user clicked No, do nothing or display a message
        MsgBox "You are still logged in.", vbInformation, "Logout Cancelled"
    End If
End Sub


Private Sub Command2_Click()

    ' Ask the user if they want to log out
    Dim response As VbMsgBoxResult
    response = MsgBox("Do you really want to Phone?", vbYesNo + vbQuestion, "Logout Confirmation")

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

Private Sub Logout_Click()

    ' Ask the user if they want to log out
    Dim response As VbMsgBoxResult
    response = MsgBox("Do you really want to log out?", vbYesNo + vbQuestion, "Logout Confirmation")

    ' Check the user's response
    If response = vbYes Then
        ' Show the Welcome form if the user clicked Yes
        Index.Show
        ' Optionally, hide or close the current form
        Me.Hide ' or Me.Close if you want to completely close the current form
    Else
        ' If the user clicked No, do nothing or display a message
        MsgBox "You are still logged in.", vbInformation, "Logout Cancelled"
    End If
End Sub



