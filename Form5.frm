VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Puntajes"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2865
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   2865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Restablecer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1390
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1390
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5318
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   8421504
      BackColor       =   16777215
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim midir As String

Private Sub Command1_Click()
Form5.Visible = False
End Sub


Private Sub Command2_Click()
If MsgBox(" ¡Restablecer marcadores? ", vbOKCancel + vbExclamation, " Eliminar marcadores") = vbOK Then
sql = "delete * from  puntuaciones"
midir = App.Path & "\recursos\db.mdb"
Call Cargar_ListView(ListView1, Trim$(sql), Trim$(midir))

Call cargar_list(Form5.ListView1)

End If

End Sub

Private Sub Form_Load()
Call cargar_list(Form5.ListView1)
End Sub

Private Sub Form_LostFocus()
If Form5.Visible = True Then Form5.SetFocus
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.BackColor = &HFFFFFF
Command2.BackColor = &HFFFFFF
End Sub

Private Sub ListView1_LostFocus()
If Form5.Visible = True Then Form5.SetFocus

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.BackColor = &HFF8080
Command2.BackColor = &HFFFFFF
Command1.SetFocus
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command2.BackColor = &HC0C000
Command1.BackColor = &HFFFFFF
Command2.SetFocus
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.BackColor = &HFFFFFF
Command2.BackColor = &HFFFFFF
End Sub
