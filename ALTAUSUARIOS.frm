VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "ALTA USUARIOS "
   ClientHeight    =   5340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7590
   LinkTopic       =   "Form2"
   ScaleHeight     =   5340
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "REGRESAR MENU"
      Height          =   975
      Left            =   5040
      TabIndex        =   7
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LIMPIAR"
      Height          =   975
      Left            =   2760
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DAR DE ALTA"
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   3240
      TabIndex        =   3
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   3240
      TabIndex        =   2
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "DAR DE ALTA A USUARIOS"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   6615
   End
   Begin VB.Label Label2 
      Caption         =   "USUARIO"
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim id As String
    Dim nombre As String
    id = Text1.Text
    nombre = Text2.Text
    
    Dim fileName As String
    fileName = App.Path & "\" & "Usuarios.dat"
    
    Open fileName For Append As #1
    Print #1, id & ";" & nombre
    Close #1
    
    MsgBox "USUARIO DADO DE ALTA CORRECTAMENTE"
    
    Text1.Text = ""
    Text2.Text = ""
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Command3_Click()
Form1.Show
Form2.Hide
End Sub
