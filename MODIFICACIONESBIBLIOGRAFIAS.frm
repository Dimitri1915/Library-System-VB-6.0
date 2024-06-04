VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "MODIFICACIONES DE BIBLIOGRAFIAS"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7650
   LinkTopic       =   "Form8"
   ScaleHeight     =   6750
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "MODIFICACIONES DE BIBLIOGRAFIAS."
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ModificarBibliografia()
    Dim isbnAModificar As String
    isbnAModificar = Text1.Text
    
    Dim nuevoTitulo As String
    nuevoTitulo = Text2.Text
    
    Dim nuevoAutor As String
    nuevoAutor = Text3.Text
    
    Dim nuevoStatus As String
    nuevoStatus = Text4.Text
    
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
        If isbn = isbnAModificar Then
            encontrado = True
            Print #2, isbn & ";" & nuevoTitulo & ";" & nuevoAutor & ";" & nuevoStatus
        Else
            Print #2, linea
        End If
    Loop
    
    Close #1
    Close #2
    
    Kill fileName
    Name tempFileName As fileName
    
    If encontrado Then
        MsgBox "Bibliografía modificada correctamente."
    Else
        MsgBox "No se encontró la bibliografía."
    End If

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
End Sub
