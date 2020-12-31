VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carol E. Castro -- Modulo de aprendizaje interactivo --"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Txt_Email 
      Height          =   405
      Left            =   2400
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Txt_Name 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Correo Electronico:"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre Completo:"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Bienvenidos, porfavor ingresa tus datos para iniciar."
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oExcel As Object
Public oBook As Object
Public oSheet As Object
Public RutaReg As String

'Este devuelve la posicion actual del video
'MsgBox Wmp.Controls.currentPositionString
'Wmp.Controls.play
'reproduce video
'Wmp.URL = "D:\videos\p.mp4"



Private Sub Command1_Click()
If Txt_Name = "" Or Txt_Email = "" Then
    MsgBox "Por favor diligencia tus datos"
End If

'Crear nuevo archivo de excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
      
'agregar datos a la primer hoja en el excel
Set oSheet = oBook.Worksheets(1)

With oSheet
    .range("A1").Value = "Nombre Estudiante"
    .range("B1").Value = Txt_Name
    .range("A2").Value = "Correo Electronico"
    .range("B2").Value = Txt_Email
    .range("A3").Value = "Fecha y hora"
    .range("B3").Value = Date & " " & Time
    .range("A4").Value = "Hora de la respuesta"
    .range("B4").Value = "Pregunta"
    .range("C4").Value = "Respuesta"
    .range("D4").Value = "Puntaje"
End With


RutaReg = App.Path + "\Registro\" & Txt_Name & ".xlsx"

'Guardar el excel y cerrarlo
oBook.SaveAs RutaReg
oExcel.Quit

Form2.Visible = True
Me.Visible = False
End Sub
