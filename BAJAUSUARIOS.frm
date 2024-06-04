VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "BAJA USUARIOS"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7590
   LinkTopic       =   "Form3"
   ScaleHeight     =   3510
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "REGRESAR AL MENU"
      Height          =   855
      Left            =   5280
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LIMPIAR"
      Height          =   855
      Left            =   2760
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DAR DE BAJA"
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Label Label2 
      Caption         =   "ID"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "DAR DE BAJA A USUARIOS POR MEDIO DE LA ID"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim idABuscar As String
    idABuscar = Text1.Text
    
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
        If id = idABuscar Then
            encontrado = True
        Else
            Print #2, linea
        End If
    Loop
    
    Close #1
    Close #2
    
    Kill fileName
    Name tempFileName As fileName
    
    If encontrado Then
        MsgBox "USUARIO DADO DE BAJA CORRECTAMENTE."
    Else
        MsgBox "NO SE ENCONTRO AL USUARIO."
    End If
    
    Text1.Text = ""
End Sub

Private Sub Command2_Click()
Text1.Text = ""
End Sub

Private Sub Command3_Click()
Form1.Show
Form3.Hide
End Sub
