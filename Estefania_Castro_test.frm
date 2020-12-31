VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
   LinkTopic       =   "Form2"
   ScaleHeight     =   6840
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5655
      Left            =   6960
      TabIndex        =   0
      Top             =   600
      Width           =   3255
      Begin VB.CommandButton Command1 
         Caption         =   "Enviar respuesta"
         Height          =   495
         Left            =   840
         TabIndex        =   5
         Top             =   4920
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Respuesta 4"
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   4
         Top             =   3360
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Respuesta 3"
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   3
         Top             =   2880
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "respuesta 2"
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   2
         Top             =   2400
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "respuesta1"
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Lbl_Pregunta 
         Caption         =   "Label1"
         Height          =   1095
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer Video 
      Height          =   5280
      Left            =   360
      TabIndex        =   7
      Top             =   720
      Width           =   6300
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   11113
      _cy             =   9313
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RutaVid As String, TimeQ As String, CurrentQ As Integer, RightQ As String
Public Opt1 As String, Opt2 As String, Opt3 As String, Opt4 As String, Quest As String
Public oExcel As Object
Public oBook As Object
Public oSheet As Object
Public RutaReg As String
Public LastReg As Integer



Private Sub Command1_Click()

Dim indice As Integer
Dim Res_Seleccionada As String

While indice < Option1.Count
    If Option1(indice).Value = True Then
        Res_Seleccionada = Option1(indice).Caption
    End If
    indice = indice + 1
Wend


'Crear nuevo archivo de excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Open(FileName:=RutaReg, ReadOnly:=False)
Set oSheet = oBook.Worksheets(1)
With oSheet
    .range("A" & LastReg).Value = Time
    .range("B" & LastReg).Value = Lbl_Pregunta
    .range("C" & LastReg).Value = Res_Seleccionada
    .range("D" & LastReg).Value = "Puntaje"
End With



oBook.Save
oExcel.quit
LastReg = LastReg + 1
Video.Controls.play
Timer1.Enabled = True
CargaConfig
Limpiar
End Sub

Private Sub Form_Load()
RutaReg = Form1.RutaReg
LastReg = 5
CurrentQ = 3
CargaConfig
Video.URL = "\Videos\" & RutaVid & ".mkv"
Timer1.Enabled = True
End Sub

Sub CargaConfig()
Dim Excel As Object
Dim LibroExcel As Object
Dim HojaExcel As Object
Dim Ruta As String

Set Excel = CreateObject("Excel.Application")
Set LibroExcel = Excel.Workbooks
Excel.Visible = False
Ruta = App.Path + ("\Cuestionarios\Cuestionario.xlsx")
Set LibroExcel = Excel.Workbooks.Open(FileName:=Ruta, ReadOnly:=True)
Set HojaExcel = LibroExcel.Sheets(1)


    RutaVid = HojaExcel.range("B1").Value
    TimeQ = HojaExcel.range("A" & CurrentQ).Value
    TimeQ = HojaExcel.range("A" & CurrentQ).Value
    Quest = HojaExcel.range("B" & CurrentQ).Value
    Opt1 = HojaExcel.range("C" & CurrentQ).Value
    Opt2 = HojaExcel.range("D" & CurrentQ).Value
    Opt3 = HojaExcel.range("E" & CurrentQ).Value
    Opt4 = HojaExcel.range("F" & CurrentQ).Value
    RightQ = HojaExcel.range("G" & CurrentQ).Value
    CurrentQ = CurrentQ + 1

If TimeQ = "" Then
    MsgBox "has finalizado chau"
    End
End If
Excel.quit
End Sub

Private Sub Timer1_Timer()
Frame1.Caption = Video.Controls.currentPositionString & "  " & TimeQ
If Video.Controls.currentPositionString = TimeQ Then
    Video.Controls.pause
    Mostrar
    Timer1.Enabled = False
End If
End Sub

Sub Mostrar()
Option1(0).Caption = Opt1
Option1(1).Caption = Opt2
Option1(2).Caption = Opt3
Option1(3).Caption = Opt4
Lbl_Pregunta.Caption = Quest
End Sub


Sub Limpiar()
Dim i As Integer
While i < Option1.Count
    Option1(i).Caption = ""
    i = i + 1
Wend
Lbl_Pregunta = ""
End Sub
