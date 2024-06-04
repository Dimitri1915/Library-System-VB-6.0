VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "CONSULTAS DE BIBLIOGRAFIAS"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   LinkTopic       =   "Form9"
   ScaleHeight     =   6630
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ConsultarBibliografia()
    Dim isbnAConsultar As String
    isbnAConsultar = Text1.Text
    
    Dim fileName As String
    fileName = App.Path & "\" & "Bibliografia.dat"
    
    Dim encontrado As Boolean
    encontrado = False
    
    Dim linea As String
    Dim isbn As String
    Dim titulo As String
    Dim autor As String
    Dim status As String
    
    Open fileName For Input As #1
    
    Do While Not EOF(1)
        Line Input #1, linea
        isbn = Split(linea, ";")(0)
        titulo = Split(linea, ";")(1)
        autor = Split(linea, ";")(2)
        status = Split(linea, ";")(3)
        If isbn = isbnAConsultar Then
            encontrado = True
            MsgBox "ISBN: " & isbn & vbCrLf & "Título: " & titulo & vbCrLf & "Autor: " & autor & vbCrLf & "Estado: " & status, vbInformation, "Detalles de la Bibliografía"
            Exit Do
        End If
    Loop
    
    Close #1
    
    If Not encontrado Then
        MsgBox "No se encontró la bibliografía.", vbExclamation, "Error"
    End If
    
    ' Limpiar el campo de texto después de la consulta
    Text1.Text = ""
End Sub
