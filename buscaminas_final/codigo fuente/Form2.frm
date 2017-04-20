VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Ganó el juego"
   ClientHeight    =   6870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7230
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   458
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   1080
   End
   Begin VB.Label entrada 
      Caption         =   "0"
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSForms.TextBox Text1 
      Height          =   480
      Left            =   1560
      TabIndex        =   2
      Top             =   3240
      Width           =   2535
      VariousPropertyBits=   1686128659
      MaxLength       =   23
      Size            =   "4471;847"
      Value           =   "Introduce tu nombre:"
      SpecialEffect   =   0
      FontName        =   "Segoe Print"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Image CmdNuevo 
      Height          =   555
      Left            =   2610
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   540
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¡Felicidades, ganó este juego!"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "es:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   1800
      TabIndex        =   0
      Top             =   5160
      Width           =   2205
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   6000
      Left            =   0
      Picture         =   "Form2.frx":07A7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5880
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conex As ADODB.Connection
Dim record As ADODB.Recordset
Dim sql As String

Private Sub CmdNuevo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.Visible = True
End Sub

Private Sub Form_Activate()
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Form2.Width = 5880
Form2.Height = 6000
End Sub

Private Sub CmdNuevo_Click()
If Text1.Text = "." Or Text1.Text = "?" Or Text1.Text = "/" Or Text1.Text = "*" Or Text1.Text = "-" Or Text1.Text = "+" Or Text1.Text = "+" Or Text1.Text = "_" Or Text1.Text = ":" Or Text1.Text = ";" Or Text1.Text = "," Or Text1.Text = "¡" Or Text1.Text = "!" Or Text1.Text = "¿" Or Text1.Text = "?" Or Text1.Text = "=" Or Text1.Text = ")" Or Text1.Text = "(" Or Text1.Text = "\" Or Text1.Text = "&" Or Text1.Text = "%" Or Text1.Text = "$" Or Text1.Text = "#" Or Text1.Text = "¬" Or Text1.Text = "|" Or Text1.Text = "°" Or Text1.Text = "<" Or Text1.Text = ">" Then
MsgBox "Caracter no valido", vbCritical, "Error"
Text1.Text = ""
Exit Sub
End If

If Text1.Text = "" Then
   MsgBox "Datos incompletos", vbCritical, "Error ..."
   Exit Sub
End If

If Text1.Text = "Introduce tu nombre:" Or Text1.Text = "Introduce tu nombre" Then
MsgBox "Introduce un nombre ", vbCritical, "Datos Incompletos"
Exit Sub
End If

If Form1.Tiempo = 0 Then
MsgBox "Eres muy bueno o no has jugado." & vbCrLf & "No se guardara tu nombre." & vbCrLf & "¡Buen intento!", vbExclamation, "Atrapado"
Text1.Text = ""
Exit Sub
End If


    sql = "Insert Into puntuaciones " _
           & "(nombre,tiempo) " _
           & "Values ('" & Text1.Text & "','" & Form1.Tiempo.Caption & "'" & ")"
        Call Ejecutar_Comando(Form5.ListView1, sql)
        
    sql = "SELECT * FROM puntuaciones order by tiempo asc"
        Call Ejecutar_Comando(Form5.ListView1, sql)
                
        Text1.Text = ""
        Form2.entrada.Caption = 0
        Form2.Hide
        
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.Visible = False
Text1.SetFocus
End Sub

Private Sub Text1_Change()
If Text1.Text = "Introduce tu nombre" Then Text1.Text = ""
If Text1.Text = "Introduce tu nombre:" Then Text1.Text = ""
End Sub

Private Sub Text1_GotFocus()
If Text1.Text = "Introduce tu nombre" Then Text1.Text = ""
End Sub




Private Sub Text1_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyReturn Then
        CmdNuevo_Click
    End If
End Sub

Private Sub Timer1_Timer()

If entrada.Caption = 0 Then
Form2.Width = 390
Form2.Height = 390
entrada.Caption = 1
End If

If entrada.Caption = 1 Then

'Width = 5880
'Height = 6000

If Form2.Width < 5700 Then
Form2.Width = Form2.Width + 100
End If

If Form2.Height < 5900 Then
Form2.Height = Form2.Height + 100
Else

Form2.entrada.Caption = 2

End If
End If

End Sub
