VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "CONSULTAS DE USUARIOS"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9630
   LinkTopic       =   "Form5"
   ScaleHeight     =   3780
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "REGRESAR MENU"
      Height          =   855
      Left            =   7680
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LIMPIAR"
      Height          =   855
      Left            =   7680
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONSULTAR USUARIO"
      Height          =   855
      Left            =   7680
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "ID"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "CONSULTAR A UN USUARIO POR MEDIO DE LA ID"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim idAConsultar As String
    idAConsultar = Text1.Text
    
    Dim fileName As String
    fileName = App.Path & "\" & "Usuarios.dat"
    
    Dim encontrado As Boolean
    encontrado = False
    
    Dim linea As String
    Dim id As String
    Dim nombre As String
    
    Open fileName For Input As #1
    
    Do While Not EOF(1)
        Line Input #1, linea
        id = Split(linea, ";")(0)
        nombre = Split(linea, ";")(1)
        If id = idAConsultar Then
            encontrado = True
            MsgBox "ID: " & id & vbCrLf & "Nombre: " & nombre, vbInformation, "Detalles del Usuario"
            Exit Do
        End If
    Loop
    
    Close #1
    
    If Not encontrado Then
        MsgBox "No se encontró el usuario.", vbExclamation, "Error"
    End If
    
    Text1.Text = ""
    Text1.SetFocus
End Sub

Private Sub Command2_Click()
Text1.Text = ""
End Sub

Private Sub Command3_Click()
Form1.Show
Form5.Hide
End Sub
