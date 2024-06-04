VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "BAJAS DE BIBLIOGRAFIAS"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8130
   LinkTopic       =   "Form7"
   ScaleHeight     =   3660
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "REGRESAR AL MENU"
      Height          =   735
      Left            =   5760
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LIMPIAR"
      Height          =   735
      Left            =   3000
      TabIndex        =   4
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DAR DE BAJA"
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "ISBN"
      Height          =   855
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "BAJAS DE BIBLIOGRAFÍA POR MEDIO DEL ISBN"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim isbnABuscar As String
    isbnABuscar = Text1.Text
    
    Dim fileName As String
    fileName = App.Path & "\" & "Bibliografia.dat"
    
    Dim tempFileName As String
    tempFileName = App.Path & "\" & "TempBibliografia.dat"
    
    Dim encontrado As Boolean
    encontrado = False
    
    Open fileName For Input As #1
    Open tempFileName For Output As #2
    
    Dim linea As String
    Dim isbn As String
    Dim titulo As String
    Dim autor As String
    Dim status As String
    
    Do While Not EOF(1)
        Line Input #1, linea
        isbn = Split(linea, ";")(0)
        titulo = Split(linea, ";")(1)
        autor = Split(linea, ";")(2)
        status = Split(linea, ";")(3)
        If isbn = isbnABuscar Then
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
        MsgBox "BIBLIOGRAFIA DADO DE BAJA CORRECTAMENTE."
    Else
        MsgBox "NO SE ENCONTRO LA BIBLIOGRAFIA."
    End If
    
    Text1.Text = ""
    Text1.SetFocus
End Sub

Private Sub Command2_Click()
Text1.Text = ""
End Sub

Private Sub Command3_Click()
Form1.Show
Form7.Hide
End Sub
