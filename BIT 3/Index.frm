VERSION 5.00
Begin VB.Form Index 
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "LOGIN"
      Height          =   855
      Left            =   3840
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "REGISTER"
      Height          =   855
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "WELCOME TO OUR SHOP"
      Height          =   735
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Register.Show
End Sub

Private Sub Command2_Click()
Login.Show
        
End Sub

