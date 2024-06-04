VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MENUS"
   ClientHeight    =   3405
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "BIENVENIDO AL MENU DE BIBLIOTECA"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
   Begin VB.Menu menu 
      Caption         =   "MENUS"
      Begin VB.Menu sal 
         Caption         =   "SALIR"
      End
   End
   Begin VB.Menu usuario 
      Caption         =   "USUARIOS"
      Begin VB.Menu alta 
         Caption         =   "ALTAS"
      End
      Begin VB.Menu baja 
         Caption         =   "BAJAS"
      End
      Begin VB.Menu modificar 
         Caption         =   "MODIFICACIÓN"
      End
      Begin VB.Menu consulta 
         Caption         =   "CONSULTAS"
      End
   End
   Begin VB.Menu bibliografia 
      Caption         =   "BIBLIOGRAFÍAS"
      Begin VB.Menu alta1 
         Caption         =   "ALTAS"
      End
      Begin VB.Menu baja1 
         Caption         =   "BAJAS"
      End
      Begin VB.Menu modificar1 
         Caption         =   "MODIFICACIÓN"
      End
      Begin VB.Menu consulta1 
         Caption         =   "CONSULTAS"
      End
   End
   Begin VB.Menu prestamo 
      Caption         =   "PRESTAMOS"
      Begin VB.Menu prestamo1 
         Caption         =   "REGISTRAR PRESTAMO"
      End
      Begin VB.Menu devolución 
         Caption         =   "REGISTRAR DEVOLUCIÓN"
      End
   End
   Begin VB.Menu reporte 
      Caption         =   "REPORTES"
      Begin VB.Menu reporte1 
         Caption         =   "GENERAR REPORTES"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
Form2.Show
Form1.Hide
End Sub

Private Sub baja_Click()
Form3.Show
Form1.Hide
End Sub

Private Sub sal_Click()
End
End Sub
