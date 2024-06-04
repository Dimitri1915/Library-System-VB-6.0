VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "MODIFICACIONES DE USUARIOS"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12615
   LinkTopic       =   "Form4"
   ScaleHeight     =   4005
   ScaleWidth      =   12615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "REGRESAR AL MENU"
      Height          =   1095
      Left            =   9600
      TabIndex        =   7
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LIMPIAR"
      Height          =   1095
      Left            =   9600
      TabIndex        =   6
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MODIFICAR USUARIO"
      Height          =   1095
      Left            =   9600
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   2160
      TabIndex        =   4
      Top             =   2880
      Width           =   7095
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   2160
      TabIndex        =   3
      Top             =   1560
      Width           =   7095
   End
   Begin VB.Label Label3 
      Caption         =   "NOMBRE"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "ID"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "MODIFICACIONES DE NOMBRE DE USUARIO POR MEDIO DE LA ID"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim idAModificar As String
    idAModificar = Text1.Text
    
    Dim nuevoNombre As String
    nuevoNombre = Text2.Text
    
    Dim fileName As String
    fileName = App.Path & "\" & "Usuarios.dat"
    
    Dim tempFileName As String
    tempFileName = App.Path & "\" & "TempUsuarios.dat"
    
    Dim encontrado As Boolean
    encontrado = False
    
    Open fileName For Input As #1
    Open tempFileName For Output As #2
    
    Dim linea As String
    Dim id As String
    Dim nombre As String
    
    Do While Not EOF(1)
        Line Input #1, linea
        id = Split(linea, ";")(0)
        nombre = Split(linea, ";")(1)
        If id = idAModificar Then
            encontrado = True
            Print #2, id & ";" & nuevoNombre
        Else
            Print #2, linea
        End If
    Loop
    
    Close #1
    Close #2
    
    Kill fileName
    Name tempFileName As fileName
    
    If encontrado Then
        MsgBox "USUARIO MODIFICADO CORRECTAMENTE."
    Else
        MsgBox "NO SE ENCONTRO AL USUARIO."
    End If
    
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Command3_Click()
Form1.Show
Form4.Hide
End Sub
