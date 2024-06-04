VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "ALTA DE BIBLIOGRAFÍAS"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8355
   LinkTopic       =   "Form6"
   ScaleHeight     =   6330
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "REGRESAR AL MENU"
      Height          =   735
      Left            =   5760
      TabIndex        =   11
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LIMPIAR"
      Height          =   735
      Left            =   3120
      TabIndex        =   10
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DAR DE ALTA"
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   1680
      TabIndex        =   8
      Top             =   4320
      Width           =   6375
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   1680
      TabIndex        =   7
      Top             =   3360
      Width           =   6375
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1680
      TabIndex        =   5
      Top             =   1440
      Width           =   6375
   End
   Begin VB.Label Label2 
      Caption         =   "STATUS"
      Height          =   615
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "AUTOR"
      Height          =   615
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "TITULO"
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "ISBN"
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ALTA DE BIBLIOGRAFÍAS "
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim isbn As String
    Dim titulo As String
    Dim autor As String
    Dim status As String
    
    isbn = Text1.Text
    titulo = Text2.Text
    autor = Text3.Text
    status = Text4.Text
    
    Dim fileName As String
    fileName = App.Path & "\" & "Bibliografia.dat"
    
    Open fileName For Append As #1
    Print #1, isbn & ";" & titulo & ";" & autor & ";" & status
    Close #1
    
    MsgBox "BIBLIOGRAFIA AGREGADA CORRECTAMENTE."
    
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text1.SetFocus
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
End Sub

Private Sub Command3_Click()
Form1.Show
Form6.Hide
End Sub

