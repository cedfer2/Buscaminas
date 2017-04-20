VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H000040C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscaminas"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   19110
   ForeColor       =   &H00C00000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1274
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   139
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   138
      Top             =   3480
      Width           =   855
   End
   Begin VB.Timer delay2 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   10560
      Top             =   240
   End
   Begin VB.Timer delay1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   10200
      Top             =   240
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   9840
      Top             =   240
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   9480
      Top             =   240
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   9120
      Top             =   240
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   6720
      TabIndex        =   113
      Top             =   2160
      Visible         =   0   'False
      Width           =   4575
      Begin VB.Image vist5 
         Height          =   1005
         Left            =   2640
         Picture         =   "Form1.frx":DB20
         Stretch         =   -1  'True
         ToolTipText     =   "Partida jugada"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image vist4 
         Height          =   1005
         Left            =   840
         Picture         =   "Form1.frx":EAF9
         Stretch         =   -1  'True
         ToolTipText     =   "Partida sin jugar"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label nota1 
         BackStyle       =   0  'Transparent
         Caption         =   "(Si no está seguro, vuelva a hacer clic con el botón secundario para desmarcarlo)."
         Height          =   855
         Left            =   2760
         TabIndex        =   124
         Top             =   2040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Image vist3 
         Height          =   1005
         Left            =   1560
         Picture         =   "Form1.frx":1041A
         Stretch         =   -1  'True
         ToolTipText     =   "Mina marcada"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image vist2 
         Height          =   1005
         Left            =   240
         Picture         =   "Form1.frx":1101A
         Stretch         =   -1  'True
         ToolTipText     =   "Mina sin marcar"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image vist1 
         Height          =   1455
         Left            =   3000
         Picture         =   "Form1.frx":11BC5
         Stretch         =   -1  'True
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Image anterior 
         Height          =   195
         Left            =   3960
         Picture         =   "Form1.frx":1416B
         Stretch         =   -1  'True
         ToolTipText     =   "Anterior"
         Top             =   3000
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image siguiente 
         Height          =   195
         Left            =   4200
         Picture         =   "Form1.frx":14268
         Stretch         =   -1  'True
         ToolTipText     =   "Siguiente"
         Top             =   3000
         Width           =   195
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   3960
         Picture         =   "Form1.frx":14362
         ToolTipText     =   "Cerrar"
         Top             =   120
         Width           =   495
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   120
         Top             =   120
         Width           =   3855
      End
      Begin VB.Shape Shape2 
         Height          =   2055
         Left            =   240
         Top             =   120
         Width           =   4560
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Las minas están ocultas debajo de los cuadros. !Cuidado! Puede perder tras un solo clic."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   21
         Left            =   120
         TabIndex        =   116
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevo juego"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   11
         Left            =   120
         TabIndex        =   115
         Top             =   360
         Width           =   2055
      End
      Begin VB.Image Image5 
         Height          =   960
         Left            =   480
         Picture         =   "Form1.frx":147A2
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   960
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Puede cambiar la apariencia del tablero y las minas en ""Cambiar de apariencia""."
         Height          =   855
         Index           =   5
         Left            =   1680
         TabIndex        =   114
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Image Image6 
         Height          =   11520
         Left            =   240
         Picture         =   "Form1.frx":17325
         Top             =   240
         Width           =   20490
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   9120
      TabIndex        =   120
      Top             =   1200
      Width           =   495
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6840
      Top             =   240
   End
   Begin VB.Timer animar 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8760
      Top             =   240
   End
   Begin VB.Timer delay 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   7320
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   8280
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7800
      Top             =   240
   End
   Begin VB.CommandButton a1 
      BackColor       =   &H0080C0FF&
      Caption         =   "1"
      Height          =   480
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2025
      TabIndex        =   127
      Top             =   4500
      Visible         =   0   'False
      Width           =   1095
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "??"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   540
         TabIndex        =   128
         ToolTipText     =   "Liberar mina"
         Top             =   120
         Width           =   270
      End
      Begin VB.Image Image8 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Left            =   0
         ToolTipText     =   "Liberar mina"
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.CommandButton a15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "15"
      Height          =   480
      Left            =   16215
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   136
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "16"
      Height          =   480
      Left            =   16695
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   135
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a42 
      BackColor       =   &H00C0C0C0&
      Caption         =   "42"
      Height          =   480
      Left            =   16215
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   132
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2040
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a44 
      BackColor       =   &H00C0C0C0&
      Caption         =   "44"
      Height          =   480
      Left            =   17175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   131
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2040
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a52 
      BackColor       =   &H00C0C0C0&
      Caption         =   "52"
      Height          =   480
      Left            =   16695
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   130
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2520
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a53 
      BackColor       =   &H00C0C0C0&
      Caption         =   "53"
      Height          =   480
      Left            =   17175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2520
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a51 
      BackColor       =   &H00C0C0C0&
      Caption         =   "51"
      Height          =   480
      Left            =   16215
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2520
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a50 
      BackColor       =   &H00C0C0C0&
      Caption         =   "50"
      Height          =   480
      Left            =   15735
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   43
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2520
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a49 
      BackColor       =   &H00C0C0C0&
      Caption         =   "49"
      Height          =   480
      Left            =   15255
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2520
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a48 
      BackColor       =   &H00C0C0C0&
      Caption         =   "48"
      Height          =   480
      Left            =   14775
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2520
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a47 
      BackColor       =   &H00C0C0C0&
      Caption         =   "47"
      Height          =   480
      Left            =   14295
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2520
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a46 
      BackColor       =   &H00C0C0C0&
      Caption         =   "46"
      Height          =   480
      Left            =   13815
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2520
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a45 
      BackColor       =   &H00C0C0C0&
      Caption         =   "45"
      Height          =   480
      Left            =   17655
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   38
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2040
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a43 
      BackColor       =   &H00C0C0C0&
      Caption         =   "43"
      Height          =   480
      Left            =   16695
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2040
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a41 
      BackColor       =   &H00C0C0C0&
      Caption         =   "41"
      Height          =   480
      Left            =   15735
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2040
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a40 
      BackColor       =   &H00C0C0C0&
      Caption         =   "40"
      Height          =   480
      Left            =   15255
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2040
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a39 
      BackColor       =   &H00C0C0C0&
      Caption         =   "39"
      Height          =   480
      Left            =   14775
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2040
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a38 
      BackColor       =   &H00C0C0C0&
      Caption         =   "38"
      Height          =   480
      Left            =   14295
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2040
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a37 
      BackColor       =   &H00C0C0C0&
      Caption         =   "37"
      Height          =   480
      Left            =   13815
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2040
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a35 
      BackColor       =   &H00C0C0C0&
      Caption         =   "35"
      Height          =   480
      Left            =   17175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1560
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a34 
      BackColor       =   &H00C0C0C0&
      Caption         =   "34"
      Height          =   480
      Left            =   16695
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1560
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a33 
      BackColor       =   &H00C0C0C0&
      Caption         =   "33"
      Height          =   480
      Left            =   16215
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1560
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a32 
      BackColor       =   &H00C0C0C0&
      Caption         =   "32"
      Height          =   480
      Left            =   15735
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1560
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a31 
      BackColor       =   &H00C0C0C0&
      Caption         =   "31"
      Height          =   480
      Left            =   15255
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1560
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a30 
      BackColor       =   &H00C0C0C0&
      Caption         =   "30"
      Height          =   480
      Left            =   14775
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1560
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a29 
      BackColor       =   &H00C0C0C0&
      Caption         =   "29"
      Height          =   480
      Left            =   14295
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1560
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a28 
      BackColor       =   &H00C0C0C0&
      Caption         =   "28"
      Height          =   480
      Left            =   13815
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1560
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a27 
      BackColor       =   &H00C0C0C0&
      Caption         =   "27"
      Height          =   480
      Left            =   17655
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a26 
      BackColor       =   &H00C0C0C0&
      Caption         =   "26"
      Height          =   480
      Left            =   17175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a25 
      BackColor       =   &H00C0C0C0&
      Caption         =   "25"
      Height          =   480
      Left            =   16695
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a24 
      BackColor       =   &H00C0C0C0&
      Caption         =   "24"
      Height          =   480
      Left            =   16215
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a23 
      BackColor       =   &H00C0C0C0&
      Caption         =   "23"
      Height          =   480
      Left            =   15735
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a22 
      BackColor       =   &H00C0C0C0&
      Caption         =   "22"
      Height          =   480
      Left            =   15255
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a21 
      BackColor       =   &H00C0C0C0&
      Caption         =   "21"
      Height          =   480
      Left            =   14775
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a20 
      BackColor       =   &H00C0C0C0&
      Caption         =   "20"
      Height          =   480
      Left            =   14295
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a19 
      BackColor       =   &H00C0C0C0&
      Caption         =   "19"
      Height          =   480
      Left            =   13815
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "17"
      Height          =   480
      Left            =   17175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a14 
      BackColor       =   &H00C0C0C0&
      Caption         =   "14"
      Height          =   480
      Left            =   15735
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "13"
      Height          =   480
      Left            =   15255
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "12"
      Height          =   480
      Left            =   14775
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "11"
      Height          =   480
      Left            =   14295
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "10"
      Height          =   480
      Left            =   13815
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "9"
      Height          =   480
      Left            =   17655
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "8"
      Height          =   480
      Left            =   17175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "6"
      Height          =   480
      Left            =   16215
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2"
      Height          =   480
      Left            =   14295
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "3"
      Height          =   480
      Left            =   14775
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "4"
      Height          =   480
      Left            =   15255
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "5"
      Height          =   480
      Left            =   15735
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "7"
      Height          =   480
      Left            =   16695
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a64 
      BackColor       =   &H00C0C0C0&
      Caption         =   "64"
      Height          =   480
      Left            =   13815
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   55
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a55 
      BackColor       =   &H00C0C0C0&
      Caption         =   "55"
      Height          =   480
      Left            =   13815
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   46
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a56 
      BackColor       =   &H00C0C0C0&
      Caption         =   "56"
      Height          =   480
      Left            =   14295
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   47
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a57 
      BackColor       =   &H00C0C0C0&
      Caption         =   "57"
      Height          =   480
      Left            =   14775
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a58 
      BackColor       =   &H00C0C0C0&
      Caption         =   "58"
      Height          =   480
      Left            =   15255
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   49
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a59 
      BackColor       =   &H00C0C0C0&
      Caption         =   "59"
      Height          =   480
      Left            =   15735
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a60 
      BackColor       =   &H00C0C0C0&
      Caption         =   "60"
      Height          =   480
      Left            =   16215
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   51
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a61 
      BackColor       =   &H00C0C0C0&
      Caption         =   "61"
      Height          =   480
      Left            =   16695
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a62 
      BackColor       =   &H00C0C0C0&
      Caption         =   "62"
      Height          =   480
      Left            =   17175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   53
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a63 
      BackColor       =   &H00C0C0C0&
      Caption         =   "63"
      Height          =   480
      Left            =   17655
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a65 
      BackColor       =   &H00C0C0C0&
      Caption         =   "65"
      Height          =   480
      Left            =   14295
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   56
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a66 
      BackColor       =   &H00C0C0C0&
      Caption         =   "66"
      Height          =   480
      Left            =   14775
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   57
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a67 
      BackColor       =   &H00C0C0C0&
      Caption         =   "67"
      Height          =   480
      Left            =   15255
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   58
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a68 
      BackColor       =   &H00C0C0C0&
      Caption         =   "68"
      Height          =   480
      Left            =   15735
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   59
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a69 
      BackColor       =   &H00C0C0C0&
      Caption         =   "69"
      Height          =   480
      Left            =   16215
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a70 
      BackColor       =   &H00C0C0C0&
      Caption         =   "70"
      Height          =   480
      Left            =   16695
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   61
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a71 
      BackColor       =   &H00C0C0C0&
      Caption         =   "71"
      Height          =   480
      Left            =   17175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   62
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a73 
      BackColor       =   &H00C0C0C0&
      Caption         =   "73"
      Height          =   480
      Left            =   13815
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   63
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3960
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a74 
      BackColor       =   &H00C0C0C0&
      Caption         =   "74"
      Height          =   480
      Left            =   14295
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   64
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3960
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a75 
      BackColor       =   &H00C0C0C0&
      Caption         =   "75"
      Height          =   480
      Left            =   14775
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   65
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3960
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a76 
      BackColor       =   &H00C0C0C0&
      Caption         =   "76"
      Height          =   480
      Left            =   15255
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   66
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3960
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a77 
      BackColor       =   &H00C0C0C0&
      Caption         =   "77"
      Height          =   480
      Left            =   15735
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   67
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3960
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a78 
      BackColor       =   &H00C0C0C0&
      Caption         =   "78"
      Height          =   480
      Left            =   16215
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   68
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3960
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a79 
      BackColor       =   &H00C0C0C0&
      Caption         =   "79"
      Height          =   480
      Left            =   16695
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   69
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3960
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a80 
      BackColor       =   &H00C0C0C0&
      Caption         =   "80"
      Height          =   480
      Left            =   17175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   70
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3960
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a81 
      BackColor       =   &H00C0C0C0&
      Caption         =   "81"
      Height          =   480
      Left            =   17655
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   71
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3960
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a72 
      BackColor       =   &H00C0C0C0&
      Caption         =   "72"
      Height          =   480
      Left            =   17655
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   137
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a54 
      BackColor       =   &H00C0C0C0&
      Caption         =   "54"
      Height          =   480
      Left            =   17655
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   129
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   2520
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a36 
      BackColor       =   &H00C0C0C0&
      Caption         =   "36"
      Height          =   480
      Left            =   17655
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   133
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   1560
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton a18 
      BackColor       =   &H00C0C0C0&
      Caption         =   "18"
      Height          =   480
      Left            =   17655
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   134
      TabStop         =   0   'False
      ToolTipText     =   "De clic con botón primario para descubrir contenido"
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image tr 
      Height          =   480
      Index           =   1
      Left            =   5640
      Picture         =   "Form1.frx":801D8
      ToolTipText     =   "Liberar mina"
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image tr 
      Height          =   480
      Index           =   0
      Left            =   5040
      Picture         =   "Form1.frx":806F9
      ToolTipText     =   "Liberar mina"
      Top             =   3600
      Width           =   480
   End
   Begin VB.Label bloq 
      Caption         =   "0"
      Height          =   255
      Left            =   6600
      TabIndex        =   126
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label cargando 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   345
      Left            =   1440
      TabIndex        =   125
      Top             =   4560
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   10800
      TabIndex        =   123
      Top             =   840
      Width           =   375
   End
   Begin VB.Label actl 
      Caption         =   "1"
      Height          =   255
      Left            =   10200
      TabIndex        =   122
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   4335
      Left            =   4455
      Top             =   135
      Width           =   255
   End
   Begin VB.Label pop 
      Caption         =   "0"
      Height          =   255
      Left            =   9360
      TabIndex        =   121
      Top             =   840
      Width           =   615
   End
   Begin VB.Image tabl 
      Height          =   480
      Index           =   3
      Left            =   7680
      Picture         =   "Form1.frx":80C24
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label fond 
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   255
      Left            =   8760
      TabIndex        =   119
      Top             =   840
      Width           =   375
   End
   Begin VB.Image fondo1 
      Appearance      =   0  'Flat
      Height          =   780
      Index           =   2
      Left            =   6720
      Picture         =   "Form1.frx":829D7
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image fondo1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   780
      Index           =   1
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image fondo1 
      Appearance      =   0  'Flat
      Height          =   780
      Index           =   0
      Left            =   4800
      Picture         =   "Form1.frx":8AC38
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label list 
      Caption         =   "0"
      Height          =   255
      Left            =   8400
      TabIndex        =   118
      Top             =   840
      Width           =   255
   End
   Begin VB.Label click 
      Caption         =   "0"
      Height          =   255
      Left            =   6840
      TabIndex        =   117
      Top             =   840
      Width           =   495
   End
   Begin VB.Image explo 
      Height          =   555
      Index           =   0
      Left            =   4680
      Picture         =   "Form1.frx":91AC9
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image explo 
      Height          =   555
      Index           =   12
      Left            =   11760
      Picture         =   "Form1.frx":91EBD
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image explo 
      Height          =   555
      Index           =   1
      Left            =   5280
      Picture         =   "Form1.frx":922C2
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image explo 
      Height          =   555
      Index           =   2
      Left            =   5880
      Picture         =   "Form1.frx":9239F
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image explo 
      Height          =   555
      Index           =   3
      Left            =   6480
      Picture         =   "Form1.frx":92793
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image explo 
      Height          =   555
      Index           =   4
      Left            =   7080
      Picture         =   "Form1.frx":92B84
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image explo 
      Height          =   555
      Index           =   5
      Left            =   7680
      Picture         =   "Form1.frx":92F7D
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image explo 
      Height          =   555
      Index           =   6
      Left            =   8280
      Picture         =   "Form1.frx":93362
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image explo 
      Height          =   555
      Index           =   7
      Left            =   8760
      Picture         =   "Form1.frx":9374D
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image explo 
      Height          =   555
      Index           =   8
      Left            =   9360
      Picture         =   "Form1.frx":93B92
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image explo 
      Height          =   555
      Index           =   9
      Left            =   9960
      Picture         =   "Form1.frx":94026
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image explo 
      Height          =   555
      Index           =   10
      Left            =   10560
      Picture         =   "Form1.frx":94484
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image explo 
      Height          =   555
      Index           =   11
      Left            =   11160
      Picture         =   "Form1.frx":9489F
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image bom 
      Height          =   480
      Index           =   2
      Left            =   5880
      Picture         =   "Form1.frx":94CF4
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image mina 
      Height          =   480
      Index           =   2
      Left            =   5880
      Picture         =   "Form1.frx":95275
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image mina 
      Height          =   480
      Index           =   1
      Left            =   5280
      Picture         =   "Form1.frx":957DC
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image bom 
      Height          =   480
      Index           =   1
      Left            =   5280
      MousePointer    =   1  'Arrow
      Picture         =   "Form1.frx":95D4D
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image mina 
      Height          =   480
      Index           =   0
      Left            =   4680
      Picture         =   "Form1.frx":96278
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image bom 
      Height          =   480
      Index           =   0
      Left            =   4680
      Picture         =   "Form1.frx":96834
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label nbom 
      Caption         =   "0"
      Height          =   255
      Left            =   7440
      TabIndex        =   112
      Top             =   840
      Width           =   375
   End
   Begin VB.Image ban 
      Height          =   480
      Index           =   2
      Left            =   5880
      Picture         =   "Form1.frx":96DAF
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image tabl 
      Height          =   480
      Index           =   2
      Left            =   5880
      Picture         =   "Form1.frx":98C98
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label cam 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   7920
      TabIndex        =   111
      Top             =   840
      Width           =   375
   End
   Begin VB.Image ban 
      Height          =   480
      Index           =   1
      Left            =   5280
      Picture         =   "Form1.frx":1C1663
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ban 
      Height          =   480
      Index           =   0
      Left            =   4680
      Picture         =   "Form1.frx":1C942C
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image tabl 
      Height          =   480
      Index           =   0
      Left            =   4680
      Picture         =   "Form1.frx":1C997E
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image tabl 
      Height          =   480
      Index           =   1
      Left            =   5280
      Picture         =   "Form1.frx":1D2656
      Stretch         =   -1  'True
      Top             =   135
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Tiempo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   3825
      TabIndex        =   110
      Top             =   4620
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Minas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   150
      TabIndex        =   109
      Top             =   4620
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   555
      Picture         =   "Form1.frx":1D2B1F
      Top             =   4680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   4050
      Picture         =   "Form1.frx":1D2ED1
      Top             =   4590
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Shape Shape1 
      Height          =   4335
      Left            =   120
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   6
      Left            =   3015
      Picture         =   "Form1.frx":1D3FA2
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   4
      Left            =   2535
      Picture         =   "Form1.frx":1D44B2
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   5
      Left            =   3495
      Picture         =   "Form1.frx":1D49C2
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   360
      Index           =   1
      Left            =   3150
      TabIndex        =   108
      Top             =   2130
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   360
      Index           =   3
      Left            =   3630
      TabIndex        =   107
      Top             =   2580
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   360
      Index           =   4
      Left            =   3630
      TabIndex        =   106
      Top             =   3090
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   360
      Index           =   2
      Left            =   4110
      TabIndex        =   105
      Top             =   2130
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   360
      Index           =   0
      Left            =   3630
      TabIndex        =   104
      Top             =   1170
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   3
      Left            =   3960
      Picture         =   "Form1.frx":1D4ED2
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   7
      Left            =   3975
      Picture         =   "Form1.frx":1D53E2
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   9
      Left            =   3975
      Picture         =   "Form1.frx":1D58F2
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   8
      Left            =   1095
      Picture         =   "Form1.frx":1D5E02
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   2
      Left            =   3975
      Picture         =   "Form1.frx":1D6312
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   3015
      Picture         =   "Form1.frx":1D6822
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Index           =   7
      Left            =   3150
      TabIndex        =   103
      Top             =   1650
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Index           =   5
      Left            =   3150
      TabIndex        =   102
      Top             =   1170
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Index           =   0
      Left            =   2670
      TabIndex        =   101
      Top             =   210
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Index           =   4
      Left            =   2670
      TabIndex        =   100
      Top             =   1170
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Index           =   9
      Left            =   2670
      TabIndex        =   99
      Top             =   2580
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Index           =   10
      Left            =   4110
      TabIndex        =   98
      Top             =   3090
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Index           =   8
      Left            =   3600
      TabIndex        =   97
      Top             =   1680
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Index           =   6
      Left            =   4110
      TabIndex        =   96
      Top             =   1170
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   1
      Left            =   4110
      TabIndex        =   95
      Top             =   210
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   5
      Left            =   2670
      TabIndex        =   94
      Top             =   1650
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   12
      Left            =   3150
      TabIndex        =   93
      Top             =   3090
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   11
      Left            =   2670
      TabIndex        =   92
      Top             =   3090
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   15
      Left            =   3630
      TabIndex        =   91
      Top             =   3570
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   19
      Left            =   3600
      TabIndex        =   90
      Top             =   4050
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   20
      Left            =   4110
      TabIndex        =   89
      Top             =   4050
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   14
      Left            =   1710
      TabIndex        =   88
      Top             =   3570
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   10
      Left            =   1710
      TabIndex        =   87
      Top             =   3090
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   9
      Left            =   1200
      TabIndex        =   86
      Top             =   3090
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   8
      Left            =   750
      TabIndex        =   85
      Top             =   3090
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   13
      Left            =   750
      TabIndex        =   84
      Top             =   3570
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   18
      Left            =   1710
      TabIndex        =   83
      Top             =   4050
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   17
      Left            =   1230
      TabIndex        =   82
      Top             =   4050
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   16
      Left            =   750
      TabIndex        =   81
      Top             =   4050
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   7
      Left            =   2190
      TabIndex        =   80
      Top             =   2580
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   6
      Left            =   2190
      TabIndex        =   79
      Top             =   2130
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   4
      Left            =   2190
      TabIndex        =   78
      Top             =   1650
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   3
      Left            =   2190
      TabIndex        =   77
      Top             =   1170
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   2
      Left            =   2190
      TabIndex        =   76
      Top             =   690
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Index           =   3
      Left            =   3630
      TabIndex        =   75
      Top             =   690
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Index           =   2
      Left            =   3630
      TabIndex        =   74
      Top             =   210
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Index           =   1
      Left            =   3150
      TabIndex        =   73
      Top             =   210
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   360
      Index           =   0
      Left            =   2190
      TabIndex        =   72
      Top             =   210
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   15
      X1              =   8
      X2              =   296
      Y1              =   264
      Y2              =   264
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   14
      X1              =   8
      X2              =   296
      Y1              =   232
      Y2              =   232
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   13
      X1              =   8
      X2              =   296
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   12
      X1              =   8
      X2              =   296
      Y1              =   168
      Y2              =   168
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   11
      X1              =   8
      X2              =   296
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   10
      X1              =   8
      X2              =   296
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   9
      X1              =   8
      X2              =   296
      Y1              =   72
      Y2              =   72
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   8
      X1              =   8
      X2              =   296
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   7
      X1              =   264
      X2              =   264
      Y1              =   8
      Y2              =   296
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   6
      X1              =   232
      X2              =   232
      Y1              =   8
      Y2              =   296
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   5
      X1              =   200
      X2              =   200
      Y1              =   8
      Y2              =   296
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   4
      X1              =   168
      X2              =   168
      Y1              =   8
      Y2              =   296
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   3
      X1              =   136
      X2              =   136
      Y1              =   8
      Y2              =   296
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   2
      X1              =   104
      X2              =   104
      Y1              =   8
      Y2              =   296
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   72
      X2              =   72
      Y1              =   8
      Y2              =   296
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   40
      X2              =   40
      Y1              =   8
      Y2              =   296
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   2535
      Picture         =   "Form1.frx":1D6DDE
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image fondo 
      Height          =   780
      Left            =   7080
      Picture         =   "Form1.frx":1D739A
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Menu jue 
      Caption         =   "Juego"
      Visible         =   0   'False
      Begin VB.Menu nue_je 
         Caption         =   "Nuevo juevo"
      End
      Begin VB.Menu seep 
         Caption         =   "-"
      End
      Begin VB.Menu skin 
         Caption         =   "Cambiar apariencia"
      End
      Begin VB.Menu pun 
         Caption         =   "Puntuaciones"
      End
      Begin VB.Menu separador 
         Caption         =   "-"
      End
      Begin VB.Menu ayuda 
         Caption         =   "Ayuda"
      End
      Begin VB.Menu minimizar 
         Caption         =   "Minimizar"
      End
      Begin VB.Menu sssee 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a As Integer
Dim bor As Integer
Dim explot As Integer
Dim s As Integer
Dim con As Integer
Dim minn As Integer
Dim cambioo As Integer
Dim ffon As Integer
Dim puk As Integer
'Dim son As String
Dim caco As Integer
Dim anim As Integer
Dim ree As Boolean
Dim sumo As Integer
Dim paco As Integer
Dim puchua As Integer
Dim consejo As Integer
Dim lmina As Integer

Sub bandera(contrr As Control, Button As Integer)
Dim pi As CommandButton
Set pi = contrr 'a3

If Button = 2 And bloq.Caption <> 1 Then
Call sonidoplay(App.Path & "\recursos\3.wav", sincronizado)
Frame1.Visible = False
Timer1.Enabled = True
    If pi.Picture <> ban(a).Picture Then
            If Minas.Caption > 0 Then
                With pi
                    .ToolTipText = "Mina marcada (No puede dar clic con botón primario)"
                    .Picture = ban(a).Picture
                    .Caption = ""
                 End With
                 Minas.Caption = Val(Minas.Caption) - 1
            End If
    Else
    
    pi.ToolTipText = "De clic con botón primario para descubrir contenido"
    Minas.Caption = Val(Minas.Caption) + 1
    pi.Picture = tabl(a).Picture
    End If

End If

End Sub


Sub aleatorio()
Dim max, min, w

max = 10
min = 1
w = (Int((max - min + 1) * Rnd + min))

Text1.Text = Text1.Text & w & vbCrLf


If a15.Picture = ban(a).Picture _
And a16.Picture = ban(a).Picture _
And a18.Picture = ban(a).Picture _
And a36.Picture = ban(a).Picture _
And a42.Picture = ban(a).Picture _
And a44.Picture = ban(a).Picture _
And a52.Picture = ban(a).Picture _
And a54.Picture = ban(a).Picture _
And a66.Picture = ban(a).Picture _
And a72.Picture = ban(a).Picture Then

MsgBox "Ayuda no disponible." & vbCrLf & "Logro el numero maximo de minas marcadas"
Exit Sub
End If








If w = 1 Then
If a15.Picture = ban(a).Picture Then
Call aleatorio
Text2.Text = Text2.Text & w & vbCrLf


Else
a15.Picture = ban(a).Picture
Label5.Caption = "x" & lmina
End If
End If


If w = 2 Then
If a16.Picture = ban(a).Picture Then
Call aleatorio
Text2.Text = Text2.Text & w & vbCrLf

Else
Label5.Caption = "x" & lmina
a16.Picture = ban(a).Picture
End If
End If


If w = 3 Then
If a18.Picture = ban(a).Picture Then
Call aleatorio
Text2.Text = Text2.Text & w & vbCrLf

Else
Label5.Caption = "x" & lmina
a18.Picture = ban(a).Picture
End If
End If

If w = 4 Then
If a36.Picture = ban(a).Picture Then
Call aleatorio
Text2.Text = Text2.Text & w & vbCrLf

Else
Label5.Caption = "x" & lmina
a36.Picture = ban(a).Picture
End If
End If

If w = 5 Then
If a42.Picture = ban(a).Picture Then
Call aleatorio
Text2.Text = Text2.Text & w & vbCrLf

Else
Label5.Caption = "x" & lmina
a42.Picture = ban(a).Picture
End If
End If


If w = 6 Then
If a44.Picture = ban(a).Picture Then
Call aleatorio
Text2.Text = Text2.Text & w & vbCrLf

Else
Label5.Caption = "x" & lmina
a44.Picture = ban(a).Picture
End If
End If

If w = 7 Then
If a52.Picture = ban(a).Picture Then
Call aleatorio
Text2.Text = Text2.Text & w & vbCrLf

Else
Label5.Caption = "x" & lmina
a52.Picture = ban(a).Picture
End If
End If

If w = 8 Then
If a54.Picture = ban(a).Picture Then
Call aleatorio
Text2.Text = Text2.Text & w & vbCrLf

Else
Label5.Caption = "x" & lmina
a54.Picture = ban(a).Picture
End If
End If


If w = 9 Then
If a66.Picture = ban(a).Picture Then
Call aleatorio
Text2.Text = Text2.Text & w & vbCrLf

Else

Label5.Caption = "x" & lmina
a66.Picture = ban(a).Picture
End If
End If


If w = 10 Then
If a72.Picture = ban(a).Picture Then
Call aleatorio
Text2.Text = Text2.Text & w & vbCrLf

Else
Label5.Caption = "x" & lmina
a72.Picture = ban(a).Picture
End If
End If


End Sub


Sub truco()
If Frame2.Visible = False Then
lmina = 1
Image8.Picture = tr(0).Picture
Image8.BorderStyle = 0
End If


If lmina > 0 And Frame2.Visible = True And Minas > 0 Then
Call aleatorio
Minas.Caption = Val(Minas.Caption) - 1
lmina = 0
Image8.Picture = tr(1).Picture
Else

If Minas.Caption = 0 Then
MsgBox "Liberar mina no disponible..." & vbCrLf & "¡Logró marcar el máximo de minas!", vbInformation, "Información"
End If
End If

Label5.Caption = "x" & lmina
End Sub


Sub notass()

If consejo = 0 And Frame1.Visible = True Then

anterior.Left = 3960
anterior.Visible = False
siguiente.Visible = True


Label2(11).Caption = "Nuevo juego"
Label1(21).FontSize = 12
Label1(21).Caption = "Las minas están ocultas debajo de los cuadros. !Cuidado! Puede perder tras un solo clic."
Image5.Visible = True
Label3(5).Visible = True
End If

If consejo = 1 And Frame1.Visible = True Then

anterior.Left = 3960
anterior.Visible = True
siguiente.Visible = True

Image5.Visible = True
Label3(5).Visible = True

vist2.Visible = False
vist3.Visible = False
nota1.Visible = False


Label2(11).Caption = "Objetivo:"
Label1(21).FontSize = 12
Label1(21).Caption = "Encontrar los recuadros vacíos evitando las minas. Cuanto más rápido vacíe el tablero, mejor será su puntuación."
End If




If consejo = 2 And Frame1.Visible = True Then

anterior.Left = 3960
anterior.Visible = True
siguiente.Visible = True

Image5.Visible = False
Label3(5).Visible = False

vist1.Visible = False


vist2.Visible = True
vist3.Visible = True
nota1.Visible = True

Label1(21).FontSize = 11
Label2(11).Caption = "Consejos y sugerencias."
Label1(21).Caption = "Si sospecha que un recuadro esconde una mina, haga clic con el botón secundario en él. Se agregará una marca al recuadro." '& vbCrLf & "(Si no está seguro, vuelva a hacer clic con el botón secundario para desmarcarlo)."

End If

If consejo = 3 And Frame1.Visible = True Then

anterior.Left = 3960
anterior.Visible = True
siguiente.Visible = True


Image5.Visible = True
Label3(5).Visible = True


Image5.Visible = False
Label3(5).Visible = False

vist2.Visible = False
vist3.Visible = False
nota1.Visible = False

vist4.Visible = False
vist5.Visible = False


vist1.Visible = True

Label1(21).FontSize = 12
Label2(11).Caption = "Consejos y sugerencias."
Label1(21).Caption = "Si se descubre un número," & vbCrLf & "indica el número de minas" & vbCrLf & "que hay ocultas en los" & vbCrLf & "ocho recuadros de" & vbCrLf & "alrededor."
End If

If consejo = 4 And Frame1.Visible = True Then

anterior.Left = 4200

anterior.Visible = True
siguiente.Visible = False


vist1.Visible = False

Image5.Visible = False
Label3(5).Visible = False

vist4.Visible = True
vist5.Visible = True

Label2(11).Caption = "Como ganar."

Label1(21).FontSize = 11
Label1(21).Caption = "Descubre los lugares vacios." & vbCrLf & "Descubre los números que indicán la ubicación de las minas ocultas." & vbCrLf & "Marque todas las minas ocultas."
End If

End Sub




Sub no_mostru()

For caco = 0 To 20
Label1(caco).Visible = False
Next caco

For caco = 0 To 10
Label2(caco).Visible = False
Next caco

For caco = 0 To 4
Label3(caco).Visible = False
Next caco

For caco = 0 To 9
Image1(caco).Visible = False
Next caco

End Sub



Sub mostru()

For caco = 0 To 20
Label1(caco).Visible = True
Next caco

For caco = 0 To 10
Label2(caco).Visible = True
Next caco

For caco = 0 To 4
Label3(caco).Visible = True
Next caco

For caco = 0 To 9
Image1(caco).Visible = True
Next caco

End Sub

Public Sub car()
minn = nbom.Caption
a = Form1.cam.Caption
ffon = Form1.fond.Caption
Dim c As Control

Form1.BackColor = &HFFFFFF

For Each c In Controls
If TypeOf c Is CommandButton Then
With c
.Caption = ""
.Picture = tabl(a).Picture
End With
End If
Next c



For con = 0 To 9
Image1(con).Picture = mina(minn).Picture
Next con


fondo.Visible = True
fondo.Left = 8
fondo.Top = 8
fondo.Width = 289
fondo.Height = 289

fondo.Picture = fondo1(ffon).Picture

If fond.Caption = 2 Then

For puk = 0 To 15
Line1(puk).BorderColor = &HFFFFFF '&H80000012
Next

Else

For puk = 0 To 15
Line1(puk).BorderColor = &HC0C0C0    '&H80000012
Next

End If
End Sub


Public Sub vis()

If anim = 2 Then
cargando.Visible = True

bloq.Caption = 1

Form1.pop = 1
Form1.MousePointer = 13



ree = 1
a45.Visible = ree
a44.Visible = ree
a43.Visible = ree
End If

If anim = 4 Then
cargando.Visible = False

a42.Visible = ree
a41.Visible = ree
a40.Visible = ree
End If

If anim = 6 Then
cargando.Visible = True


a39.Visible = ree
a38.Visible = ree
a37.Visible = ree
End If

If anim = 8 Then
cargando.Visible = False


a36.Visible = ree
a35.Visible = ree
a34.Visible = ree
End If

If anim = 10 Then
cargando.Visible = True

a33.Visible = ree
a32.Visible = ree
a31.Visible = ree
End If

If anim = 12 Then
cargando.Visible = False


a30.Visible = ree
a29.Visible = ree
a28.Visible = ree
End If

If anim = 14 Then
cargando.Visible = True


a27.Visible = ree
a26.Visible = ree
a25.Visible = ree
End If

If anim = 16 Then
cargando.Visible = False

a24.Visible = ree
a23.Visible = ree
a22.Visible = ree
End If

If anim = 18 Then
cargando.Visible = True

a21.Visible = ree
a20.Visible = ree
a19.Visible = ree
End If

If anim = 20 Then
cargando.Visible = False

a18.Visible = ree
a17.Visible = ree
a16.Visible = ree
End If

If anim = 22 Then
cargando.Visible = True

a15.Visible = ree
a14.Visible = ree
a13.Visible = ree
End If

If anim = 24 Then
cargando.Visible = False


a12.Visible = ree
a11.Visible = ree
a10.Visible = ree
End If

If anim = 26 Then
cargando.Visible = True

a9.Visible = ree
a8.Visible = ree
a7.Visible = ree
End If

If anim = 28 Then
cargando.Visible = False

a6.Visible = ree
a5.Visible = ree
a4.Visible = ree
End If

If anim = 30 Then
cargando.Visible = True

a3.Visible = ree
a2.Visible = ree
a1.Visible = ree
End If


If anim = 32 Then
cargando.Visible = False

a46.Visible = ree
a47.Visible = ree
a48.Visible = ree
End If


If anim = 34 Then
cargando.Visible = True

a49.Visible = ree
a50.Visible = ree
a51.Visible = ree
End If


If anim = 36 Then
cargando.Visible = False

a52.Visible = ree
a53.Visible = ree
a54.Visible = ree
End If

If anim = 38 Then
cargando.Visible = True

a55.Visible = ree
a56.Visible = ree
a57.Visible = ree
End If

If anim = 40 Then
cargando.Visible = False

a58.Visible = ree
a59.Visible = ree
a60.Visible = ree
End If

If anim = 42 Then
cargando.Visible = True

a61.Visible = ree
a62.Visible = ree
a63.Visible = ree
End If


If anim = 44 Then
cargando.Visible = False

a64.Visible = ree
a65.Visible = ree
a66.Visible = ree
End If

If anim = 46 Then
cargando.Visible = True

a67.Visible = ree
a68.Visible = ree
a69.Visible = ree
End If

If anim = 48 Then
cargando.Visible = False

a70.Visible = ree
a71.Visible = ree
a72.Visible = ree
End If


If anim = 50 Then
cargando.Visible = True

a73.Visible = ree
a74.Visible = ree
a75.Visible = ree
End If

If anim = 52 Then
cargando.Visible = False

a76.Visible = ree
a77.Visible = ree
a78.Visible = ree
End If

If anim = 54 Then
cargando.Visible = True

a79.Visible = ree
a80.Visible = ree
a81.Visible = ree

End If

If anim = 55 Then
cargando.Visible = False

Timer5.Enabled = False
anim = 0

bloq.Caption = 0
Call truco

Frame2.Visible = True

Minas.Visible = True
Image3.Visible = True
Tiempo.Visible = True
Image2.Visible = True

Form1.pop = 0
Form1.MousePointer = 0
End If

If anim Mod 2 = 0 Then Call sonidoplay(App.Path & "\recursos\6.wav", sincronizado)


End Sub

Public Sub invis()
Call no_mostru
Timer1.Enabled = False
Tiempo.Caption = 0
Frame1.Visible = False

If sumo Mod 2 = 0 Then Call sonidoplay(App.Path & "\recursos\3.wav", sincronizado)

If sumo = 2 Then
Minas.Visible = False
Image3.Visible = False
Tiempo.Visible = False
Image2.Visible = False


Frame2.Visible = False

cargando.Visible = True

bloq.Caption = 1

Form1.pop = 1
Form1.MousePointer = 13

a41.Visible = False
a42.Visible = False
End If

If sumo = 4 Then
cargando.Visible = False

a33.Visible = False
a32.Visible = False
End If

If sumo = 6 Then
cargando.Visible = True

a31.Visible = False
a40.Visible = False
End If

If sumo = 8 Then
cargando.Visible = False
a49.Visible = False
a50.Visible = False
End If

If sumo = 10 Then
cargando.Visible = True

a51.Visible = False
a52.Visible = False
End If

If sumo = 12 Then
cargando.Visible = False
a43.Visible = False
a34.Visible = False
End If

If sumo = 14 Then
cargando.Visible = True
a25.Visible = False
a24.Visible = False
End If

If sumo = 16 Then
cargando.Visible = False

a23.Visible = False
a22.Visible = False
End If

If sumo = 18 Then
cargando.Visible = True
a21.Visible = False
a30.Visible = False
End If

If sumo = 20 Then
cargando.Visible = False

a39.Visible = False
a48.Visible = False
End If

If sumo = 22 Then
cargando.Visible = True

a57.Visible = False
a58.Visible = False
End If

If sumo = 24 Then
cargando.Visible = False

a59.Visible = False
a60.Visible = False
End If

If sumo = 26 Then
cargando.Visible = True

a61.Visible = False
a62.Visible = False
End If

If sumo = 28 Then
cargando.Visible = False

a53.Visible = False
a44.Visible = False
End If


If sumo = 30 Then
cargando.Visible = True

a35.Visible = False
a26.Visible = False
End If


If sumo = 32 Then
cargando.Visible = False

a17.Visible = False
a16.Visible = False
End If


If sumo = 34 Then
cargando.Visible = True

a15.Visible = False
a14.Visible = False
End If


If sumo = 36 Then
cargando.Visible = False

a13.Visible = False
a12.Visible = False
End If


If sumo = 38 Then
cargando.Visible = True

a11.Visible = False
a20.Visible = False
End If


If sumo = 40 Then
cargando.Visible = False

a29.Visible = False
a38.Visible = False
End If


If sumo = 42 Then
cargando.Visible = True

a47.Visible = False
a56.Visible = False
End If

If sumo = 44 Then
cargando.Visible = False

a65.Visible = False
a66.Visible = False
End If

If sumo = 46 Then
cargando.Visible = True

a67.Visible = False
a68.Visible = False
End If

If sumo = 48 Then
cargando.Visible = False

a69.Visible = False
a70.Visible = False
End If

If sumo = 50 Then
cargando.Visible = True

a71.Visible = False
a72.Visible = False
End If

If sumo = 52 Then
cargando.Visible = False

a63.Visible = False
a54.Visible = False
End If

If sumo = 54 Then
cargando.Visible = True

a45.Visible = False
a36.Visible = False
End If

If sumo = 56 Then
cargando.Visible = False

a27.Visible = False
a18.Visible = False
End If

If sumo = 58 Then
cargando.Visible = True

a18.Visible = False
a9.Visible = False
End If


If sumo = 60 Then
cargando.Visible = False

a8.Visible = False
a7.Visible = False
End If

If sumo = 62 Then
cargando.Visible = True

a6.Visible = False
a5.Visible = False
End If

If sumo = 64 Then
cargando.Visible = False

a4.Visible = False
a3.Visible = False
End If


If sumo = 66 Then
cargando.Visible = True

a2.Visible = False
a1.Visible = False
End If


If sumo = 68 Then
cargando.Visible = False

a10.Visible = False
a19.Visible = False
End If

If sumo = 70 Then
cargando.Visible = True

a28.Visible = False
a37.Visible = False
End If

If sumo = 72 Then
cargando.Visible = False

a46.Visible = False
a55.Visible = False
End If


If sumo = 74 Then
cargando.Visible = True

a64.Visible = False
a73.Visible = False
End If

If sumo = 76 Then
cargando.Visible = False

a74.Visible = False
a75.Visible = False
End If

If sumo = 78 Then
cargando.Visible = True

a76.Visible = False
a77.Visible = False
End If

If sumo = 80 Then
cargando.Visible = False

a78.Visible = False
a79.Visible = False
End If


If sumo = 81 Then
cargando.Visible = True

a80.Visible = False
a81.Visible = False
End If


If sumo = 82 Then
cargando.Visible = False

Timer6.Enabled = False
sumo = 0
Form1.pop = 0
Form1.MousePointer = 0
Timer5.Enabled = True
Call nuevo
End If


End Sub


Public Sub nuevo()

'Call vis
Call car


Form1.Width = 4665
Form1.Height = 5385
Frame1.Top = 200
Frame1.Left = 0
Shape2.Left = 0
Shape2.Top = 0
Image6.Left = 0
Image6.Top = 0


Minas.Caption = 10
Tiempo.Caption = 0
Timer1.Enabled = False

Form1.list.Caption = 0

Timer2.Enabled = True
Timer4.Enabled = False
Call no_mostru
End Sub

Public Sub invocar()
Form1.pop = 1
Form1.MousePointer = 13
If bor Mod 2 = 0 Then Call sonidoplay(App.Path & "\recursos\2.wav", sincronizado)

If bor = 2 Then
Call mostru
If a1.Picture <> ban(a).Picture Then
Form1.a1.Visible = False
End If
End If

If bor = 4 And a2.Picture <> ban(a).Picture Then Form1.a2.Visible = False
If bor = 6 And a10.Picture <> ban(a).Picture Then Form1.a10.Visible = False

If bor = 8 And a3.Picture <> ban(a).Picture Then Form1.a3.Visible = False
If bor = 10 And a11.Picture <> ban(a).Picture Then Form1.a11.Visible = False
If bor = 12 And a19.Picture <> ban(a).Picture Then Form1.a19.Visible = False

If bor = 14 And a4.Picture <> ban(a).Picture Then Form1.a4.Visible = False
If bor = 16 And a12.Picture <> ban(a).Picture Then Form1.a12.Visible = False
If bor = 18 And a20.Picture <> ban(a).Picture Then Form1.a20.Visible = False
If bor = 20 And a28.Picture <> ban(a).Picture Then Form1.a28.Visible = False

If bor = 22 And a5.Picture <> ban(a).Picture Then Form1.a5.Visible = False
If bor = 24 And a13.Picture <> ban(a).Picture Then Form1.a13.Visible = False
If bor = 26 And a21.Picture <> ban(a).Picture Then Form1.a21.Visible = False
If bor = 28 And a29.Picture <> ban(a).Picture Then Form1.a29.Visible = False
If bor = 30 And a37.Picture <> ban(a).Picture Then Form1.a37.Visible = False

If bor = 32 And a14.Picture <> ban(a).Picture Then Form1.a14.Visible = False
If bor = 34 And a22.Picture <> ban(a).Picture Then Form1.a22.Visible = False
If bor = 36 And a30.Picture <> ban(a).Picture Then Form1.a30.Visible = False
If bor = 38 And a38.Picture <> ban(a).Picture Then Form1.a38.Visible = False
If bor = 40 And a46.Picture <> ban(a).Picture Then Form1.a46.Visible = False

If bor = 42 And a23.Picture <> ban(a).Picture Then Form1.a23.Visible = False
If bor = 44 And a31.Picture <> ban(a).Picture Then Form1.a31.Visible = False
If bor = 46 And a39.Picture <> ban(a).Picture Then Form1.a39.Visible = False
If bor = 48 And a47.Picture <> ban(a).Picture Then Form1.a47.Visible = False

If bor = 50 And a32.Picture <> ban(a).Picture Then Form1.a32.Visible = False
If bor = 52 And a40.Picture <> ban(a).Picture Then Form1.a40.Visible = False
If bor = 54 And a48.Picture <> ban(a).Picture Then Form1.a48.Visible = False


If bor = 56 And a41.Picture <> ban(a).Picture Then Form1.a41.Visible = False
If bor = 58 And a49.Picture <> ban(a).Picture Then Form1.a49.Visible = False

If bor = 60 Then

If a50.Picture <> ban(a).Picture Then
Form1.a50.Visible = False
End If

delay.Enabled = False
bor = 0

Form1.pop = 0
Form1.MousePointer = 0


End If

End Sub

Sub invocar1()
Form1.pop = 1
Form1.MousePointer = 13
If paco Mod 2 = 0 Then Call sonidoplay(App.Path & "\recursos\2.wav", sincronizado)

If paco = 2 Then
Call mostru

If a59.Picture <> ban(a).Picture Then
Form1.a59.Visible = False
End If
End If

If paco = 4 And a60.Picture <> ban(a).Picture Then Form1.a60.Visible = False
If paco = 6 And a68.Picture <> ban(a).Picture Then Form1.a68.Visible = False

If paco = 8 And a69.Picture <> ban(a).Picture Then Form1.a69.Visible = False
If paco = 10 And a77.Picture <> ban(a).Picture Then Form1.a77.Visible = False

If paco = 12 And a70.Picture <> ban(a).Picture Then Form1.a70.Visible = False
If paco = 14 And a78.Picture <> ban(a).Picture Then Form1.a78.Visible = False

If paco = 16 Then

If a79.Picture <> ban(a).Picture Then
Form1.a79.Visible = False
End If

delay1.Enabled = False
paco = 0

Form1.pop = 0
Form1.MousePointer = 0
End If

End Sub

Sub invocar2()

If puchua = 2 Then
Call mostru

If a55.Picture <> ban(a).Picture Then
Form1.a55.Visible = False
Call sonidoplay(App.Path & "\recursos\2.wav", sincronizado)
End If

Form1.pop = 1
Form1.MousePointer = 13

End If

If puchua = 4 And a64.Picture <> ban(a).Picture Then
Form1.a64.Visible = False
Call sonidoplay(App.Path & "\recursos\2.wav", sincronizado)
End If

If puchua = 6 Then

If a73.Picture <> ban(a).Picture Then
Call sonidoplay(App.Path & "\recursos\2.wav", sincronizado)
Form1.a73.Visible = False
End If

delay2.Enabled = False
puchua = 0

Form1.pop = 0
Form1.MousePointer = 0
End If


End Sub









Public Sub perder()
Timer1.Enabled = False
s = nbom.Caption
Form1.a18.Visible = False
Form1.a72.Visible = False
Form1.a54.Visible = False
Form1.a36.Visible = False
Form1.a18.Visible = False
Form1.a16.Visible = False
Form1.a15.Visible = False
Form1.a44.Visible = False
Form1.a52.Visible = False
Form1.a42.Visible = False
Form1.a66.Visible = False
animar.Enabled = True
'--------------------------------------
If explot = 0 Then
For con = 0 To 9
Image1(con).Picture = Nothing
Next con
End If

If explot = 1 Then
bloq.Caption = 1

Image1(0).Picture = explo(0).Picture
Image1(1).Picture = explo(0).Picture
Image1(2).Picture = explo(0).Picture
Image1(3).Picture = explo(0).Picture
Image1(4).Picture = explo(0).Picture
Image1(5).Picture = explo(0).Picture
Image1(6).Picture = explo(0).Picture
Image1(7).Picture = explo(0).Picture
Image1(8).Picture = explo(0).Picture
Image1(9).Picture = explo(0).Picture
End If

If explot = 2 Then
Image1(0).Picture = explo(1).Picture
Image1(1).Picture = explo(1).Picture
Image1(2).Picture = explo(1).Picture
Image1(3).Picture = explo(1).Picture
Image1(4).Picture = explo(1).Picture
Image1(5).Picture = explo(1).Picture
Image1(6).Picture = explo(1).Picture
Image1(7).Picture = explo(1).Picture
Image1(8).Picture = explo(1).Picture
Image1(9).Picture = explo(1).Picture

End If

If explot = 3 Then
Image1(0).Picture = explo(2).Picture
Image1(1).Picture = explo(2).Picture
Image1(2).Picture = explo(2).Picture
Image1(3).Picture = explo(2).Picture
Image1(4).Picture = explo(2).Picture
Image1(5).Picture = explo(2).Picture
Image1(6).Picture = explo(2).Picture
Image1(7).Picture = explo(2).Picture
Image1(8).Picture = explo(2).Picture
Image1(9).Picture = explo(2).Picture
End If

If explot = 4 Then
Image1(0).Picture = explo(3).Picture
Image1(1).Picture = explo(3).Picture
Image1(2).Picture = explo(3).Picture
Image1(3).Picture = explo(3).Picture
Image1(4).Picture = explo(3).Picture
Image1(5).Picture = explo(3).Picture
Image1(6).Picture = explo(3).Picture
Image1(7).Picture = explo(3).Picture
Image1(8).Picture = explo(3).Picture
Image1(9).Picture = explo(3).Picture
End If

If explot = 5 Then
Image1(0).Picture = explo(4).Picture
Image1(1).Picture = explo(4).Picture
Image1(2).Picture = explo(4).Picture
Image1(3).Picture = explo(4).Picture
Image1(4).Picture = explo(4).Picture
Image1(5).Picture = explo(4).Picture
Image1(6).Picture = explo(4).Picture
Image1(7).Picture = explo(4).Picture
Image1(8).Picture = explo(4).Picture
Image1(9).Picture = explo(4).Picture
End If

If explot = 6 Then
Image1(0).Picture = explo(6).Picture
Image1(1).Picture = explo(6).Picture
Image1(2).Picture = explo(6).Picture
Image1(3).Picture = explo(6).Picture
Image1(4).Picture = explo(6).Picture
Image1(5).Picture = explo(6).Picture
Image1(6).Picture = explo(6).Picture
Image1(7).Picture = explo(6).Picture
Image1(8).Picture = explo(6).Picture
Image1(9).Picture = explo(6).Picture
End If

If explot = 7 Then
Image1(0).Picture = explo(7).Picture
Image1(1).Picture = explo(7).Picture
Image1(2).Picture = explo(7).Picture
Image1(3).Picture = explo(7).Picture
Image1(4).Picture = explo(7).Picture
Image1(5).Picture = explo(7).Picture
Image1(6).Picture = explo(7).Picture
Image1(7).Picture = explo(7).Picture
Image1(8).Picture = explo(7).Picture
Image1(9).Picture = explo(7).Picture
End If

If explot = 8 Then
Call sonidoplay(App.Path & "\recursos\4.wav", sincronizado)
Image1(0).Picture = explo(8).Picture
Image1(1).Picture = explo(8).Picture
Image1(2).Picture = explo(8).Picture
Image1(3).Picture = explo(8).Picture
Image1(4).Picture = explo(8).Picture
Image1(5).Picture = explo(8).Picture
Image1(6).Picture = explo(8).Picture
Image1(7).Picture = explo(8).Picture
Image1(8).Picture = explo(8).Picture
Image1(9).Picture = explo(8).Picture
End If

If explot = 9 Then
Image1(0).Picture = explo(9).Picture
Image1(1).Picture = explo(9).Picture
Image1(2).Picture = explo(9).Picture
Image1(3).Picture = explo(9).Picture
Image1(4).Picture = explo(9).Picture
Image1(5).Picture = explo(9).Picture
Image1(6).Picture = explo(9).Picture
Image1(7).Picture = explo(9).Picture
Image1(8).Picture = explo(9).Picture
Image1(9).Picture = explo(9).Picture
End If

If explot = 10 Then
Image1(0).Picture = explo(10).Picture
Image1(1).Picture = explo(10).Picture
Image1(2).Picture = explo(10).Picture
Image1(3).Picture = explo(10).Picture
Image1(4).Picture = explo(10).Picture
Image1(5).Picture = explo(10).Picture
Image1(6).Picture = explo(10).Picture
Image1(7).Picture = explo(10).Picture
Image1(8).Picture = explo(10).Picture
Image1(9).Picture = explo(10).Picture
End If

If explot = 11 Then
Image1(0).Picture = explo(11).Picture
Image1(1).Picture = explo(11).Picture
Image1(2).Picture = explo(11).Picture
Image1(3).Picture = explo(11).Picture
Image1(4).Picture = explo(11).Picture
Image1(5).Picture = explo(11).Picture
Image1(6).Picture = explo(11).Picture
Image1(7).Picture = explo(11).Picture
Image1(8).Picture = explo(11).Picture
Image1(9).Picture = explo(11).Picture
End If

If explot = 12 Then
Image1(0).Picture = explo(12).Picture
Image1(1).Picture = explo(12).Picture
Image1(2).Picture = explo(12).Picture
Image1(3).Picture = explo(12).Picture
Image1(4).Picture = explo(12).Picture
Image1(5).Picture = explo(12).Picture
Image1(6).Picture = explo(12).Picture
Image1(7).Picture = explo(12).Picture
Image1(8).Picture = explo(12).Picture
Image1(9).Picture = explo(12).Picture
End If

If explot = 15 Then

For con = 0 To 9
Image1(con).Picture = mina(s).Picture
Next con

animar.Enabled = False

If click.Caption = 0 Then Image1(0).Picture = bom(s).Picture
If click.Caption = 1 Then Image1(1).Picture = bom(s).Picture
If click.Caption = 2 Then Image1(2).Picture = bom(s).Picture
If click.Caption = 3 Then Image1(3).Picture = bom(s).Picture
If click.Caption = 4 Then Image1(4).Picture = bom(s).Picture
If click.Caption = 5 Then Image1(5).Picture = bom(s).Picture
If click.Caption = 6 Then Image1(6).Picture = bom(s).Picture
If click.Caption = 7 Then Image1(7).Picture = bom(s).Picture
If click.Caption = 8 Then Image1(8).Picture = bom(s).Picture
If click.Caption = 9 Then Image1(9).Picture = bom(s).Picture
explot = 0
Form4.Timer1.Enabled = True
Form4.Show 1

Form6.Label2.Caption = "Tiempo: " & Form1.Tiempo.Caption & " segundos"
Form6.Show 1

Timer6.Enabled = True

End If
End Sub

Sub experimental(conk As Control, tipo_b As Integer)
'0 = vacio
'1 = bomba
'2 = invocar
'3 = invocar1
'4 = invocar2

Dim experi As CommandButton
Set experi = conk


If experi.Picture <> ban(a).Picture And bloq.Caption <> 1 Then
Frame1.Visible = False
Timer1.Enabled = True
Call sonidoplay(App.Path & "\recursos\1.wav", sincronizado)

If tipo_b = 0 Then
Call mostru
experi.Visible = False
End If

If tipo_b = 1 Then
Call mostru
experi.Visible = False
Call perder
click.Caption = 2
End If

If tipo_b = 2 Then
experi.Visible = False
delay.Enabled = True
End If

If tipo_b = 3 Then
experi.Visible = False
delay2.Enabled = True
End If


If tipo_b = 4 Then
experi.Visible = False
delay1.Enabled = True
End If

End If
End Sub





Private Sub animar_Timer()
explot = explot + 1
Call perder
End Sub


Private Sub anterior_Click()
Call sonidoplay(App.Path & "\recursos\2.wav", sincronizado)
consejo = consejo - 1
Call notass
End Sub

Private Sub ayuda_Click()
Timer4.Enabled = True
Frame1.Visible = True
End Sub

Private Sub cam_Change()
Timer6.Enabled = True
'Call nuevo
End Sub


Private Sub a1_Click()
Call experimental(a1, 2)
End Sub

Private Sub a2_Click()
Call experimental(a2, 2)
End Sub

Private Sub a3_Click()
Call experimental(a3, 2)
End Sub

Private Sub a4_Click()
Call experimental(a4, 2)
End Sub

Private Sub a10_Click()
Call experimental(a10, 2)
End Sub

Private Sub a11_Click()
Call experimental(a11, 2)
End Sub

Private Sub a12_Click()
Call experimental(a12, 2)
End Sub

Private Sub a13_Click()
Call experimental(a13, 2)
End Sub

Private Sub a19_Click()
Call experimental(a19, 2)
End Sub

Private Sub a20_Click()
Call experimental(a20, 2)
End Sub

Private Sub a21_Click()
Call experimental(a21, 2)
End Sub

Private Sub a22_Click()
Call experimental(a22, 2)
End Sub

Private Sub a28_Click()
Call experimental(a28, 2)
End Sub

Private Sub a29_Click()
Call experimental(a29, 2)
End Sub

Private Sub a30_Click()
Call experimental(a30, 2)
End Sub

Private Sub a31_Click()
Call experimental(a31, 2)
End Sub

Private Sub a37_Click()
Call experimental(a37, 2)
End Sub

Private Sub a38_Click()
Call experimental(a38, 2)
End Sub

Private Sub a39_Click()
Call experimental(a39, 2)
End Sub

Private Sub a40_Click()
Call experimental(a40, 2)
End Sub

Private Sub a46_Click()
Call experimental(a46, 2)
End Sub

Private Sub a47_Click()
Call experimental(a47, 2)
End Sub

Private Sub a48_Click()
Call experimental(a48, 2)
End Sub

Private Sub a49_Click()
Call experimental(a49, 2)
End Sub




Private Sub a15_Click()
Call experimental(a15, 1)
End Sub

Private Sub a16_Click()
Call experimental(a16, 1)
End Sub

Private Sub a18_Click()
Call experimental(a18, 1)
End Sub

Private Sub a36_Click()
Call experimental(a36, 1)
End Sub

Private Sub a42_Click()
Call experimental(a44, 1)
End Sub

Private Sub a44_Click()
Call experimental(a44, 1)
End Sub

Private Sub a52_Click()
Call experimental(a54, 1)
End Sub

Private Sub a54_Click()
Call experimental(a54, 1)
End Sub

Private Sub a66_Click()
Call experimental(a72, 1)
End Sub

Private Sub a72_Click()
Call experimental(a72, 1)
End Sub




Private Sub a73_Click()
Call experimental(a73, 3)
End Sub

Private Sub a64_Click()
Call experimental(a64, 3)
End Sub

Private Sub a55_Click()
Call experimental(a55, 3)
End Sub





Private Sub a79_Click()
Call experimental(a79, 4)
End Sub

Private Sub a78_Click()
Call experimental(a78, 4)
End Sub

Private Sub a77_Click()
Call experimental(a77, 4)
End Sub

Private Sub a68_Click()
Call experimental(a68, 4)
End Sub

Private Sub a69_Click()
Call experimental(a69, 4)
End Sub

Private Sub a70_Click()
Call experimental(a70, 4)
End Sub

Private Sub a59_Click()
Call experimental(a59, 4)
End Sub




Private Sub a5_Click()
Call experimental(a5, 0)
End Sub

Private Sub a6_Click()
Call experimental(a6, 0)
End Sub

Private Sub a7_Click()
Call experimental(a7, 0)
End Sub

Private Sub a8_Click()
Call experimental(a8, 0)
End Sub

Private Sub a9_Click()
Call experimental(a9, 0)
End Sub

Private Sub a14_Click()
Call experimental(a14, 0)
End Sub

Private Sub a17_Click()
Call experimental(a17, 0)
End Sub

Private Sub a23_Click()
Call experimental(a23, 0)
End Sub


Private Sub a24_Click()
Call experimental(a24, 0)
End Sub

Private Sub a25_Click()
Call experimental(a25, 0)
End Sub

Private Sub a26_Click()
Call experimental(a26, 0)
End Sub

Private Sub a27_Click()
Call experimental(a27, 0)
End Sub

Private Sub a32_Click()
Call experimental(a32, 0)
End Sub

Private Sub a33_Click()
Call experimental(a33, 0)
End Sub

Private Sub a34_Click()
Call experimental(a34, 0)
End Sub

Private Sub a35_Click()
Call experimental(a35, 0)
End Sub

Private Sub a41_Click()
Call experimental(a41, 0)
End Sub

Private Sub a43_Click()
Call experimental(a43, 0)
End Sub

Private Sub a45_Click()
Call experimental(a45, 0)
End Sub

Private Sub a50_Click()
Call experimental(a50, 0)
End Sub

Private Sub a51_Click()
Call experimental(a51, 0)
End Sub

Private Sub a53_Click()
Call experimental(a53, 0)
End Sub

Private Sub a60_Click()
Call experimental(a60, 0)
End Sub

Private Sub a61_Click()
Call experimental(a61, 0)
End Sub

Private Sub a62_Click()
Call experimental(a62, 0)
End Sub

Private Sub a63_Click()
Call experimental(a63, 0)
End Sub


Private Sub a71_Click()
Call experimental(a71, 0)
End Sub

Private Sub a80_Click()
Call experimental(a80, 0)
End Sub

Private Sub a81_Click()
Call experimental(a81, 0)
End Sub

Private Sub a56_Click()
Call experimental(a56, 0)
End Sub

Private Sub a57_Click()
Call experimental(a57, 0)
End Sub

Private Sub a58_Click()
Call experimental(a58, 0)
End Sub

Private Sub a65_Click()
Call experimental(a65, 0)
End Sub

Private Sub a67_Click()
Call experimental(a67, 0)
End Sub

Private Sub a74_Click()
Call experimental(a74, 0)
End Sub

Private Sub a75_Click()
Call experimental(a75, 0)
End Sub

Private Sub a76_Click()
Call experimental(a76, 0)
End Sub


Private Sub delay_Timer()
bor = bor + 1
Call invocar
End Sub

Private Sub delay1_Timer()
paco = paco + 1
Call invocar1

End Sub

Private Sub delay2_Timer()
puchua = puchua + 1
Call invocar2
End Sub


Private Sub fond_Change()
'Call nuevo
Timer6.Enabled = True

End Sub

Private Sub fondo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pop = 0 Then
If Button = 2 Then
PopupMenu jue
End If
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pop = 0 Then
If Button = 2 Then
PopupMenu jue
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call sonidoplay(App.Path & "\recursos\7.wav", sincronizado)
End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer4.Enabled = True
If list.Caption = 0 And Frame1.Top = 117 Then
list.Caption = 1
End If

If list.Caption = 1 And Frame1.Top = 201 Then
list.Caption = 0
End If
End Sub

Private Sub Image8_Click()
Call sonidoplay(App.Path & "\recursos\2.wav", sincronizado)
Call truco
End Sub

Private Sub Label5_Click()
Call sonidoplay(App.Path & "\recursos\2.wav", sincronizado)
Call truco
End Sub

Private Sub minimizar_Click()
Form1.WindowState = 1
End Sub

Private Sub nbom_Change()
'Call nuevo
Timer6.Enabled = True
End Sub


Private Sub nue_je_Click()
'Call nuevo
Frame1.Visible = False
Timer6.Enabled = True
End Sub

Private Sub pun_Click()
Form5.Show 1
End Sub

Private Sub salir_Click()
End
End Sub

Private Sub siguiente_Click()
Call sonidoplay(App.Path & "\recursos\2.wav", sincronizado)
consejo = consejo + 1
Call notass
End Sub

Private Sub skin_Click()
Dim ge As String
ge = MsgBox("Si cambia de apariencia se reiniciara la partida", vbExclamation, "Advertencia")
Form3.Show 1
End Sub

Private Sub Timer1_Timer()
Tiempo.Caption = Val(Tiempo.Caption) + 1
End Sub

Private Sub Timer2_Timer()
Dim b As String
Dim z As String
z = False
b = ban(a).Picture


If a66.Picture = b And a72.Picture = b And a54.Picture = b And a52.Picture = b And a44.Picture = b And a42.Picture = b And a36.Picture = b And a18.Picture = b And a16.Picture = b And a15.Picture = b _
And a5.Visible = z And a6.Visible = z And a7.Visible = z And a8.Visible = z And a9.Visible = z And a17.Visible = z And a14.Visible = z And a27.Visible = z And a26.Visible = z And a25.Visible = z And a24.Visible = z And a23.Visible = z And a35.Visible = z And a34.Visible = z And a33.Visible = z And a32.Visible = z And a45.Visible = z And a43.Visible = z And a41.Visible = z And a53.Visible = z And a51.Visible = z And a50.Visible = z And a63.Visible = z And a62.Visible = z And a61.Visible = z And a60.Visible = z And a59.Visible = z And a58.Visible = z And a57.Visible = z And a56.Visible = z And a55.Visible = z And a71.Visible = z And a70.Visible = z And a69.Visible = z And a68.Visible = z And a67.Visible = z And a65.Visible = z And a64.Visible = z And a81.Visible = z And a80.Visible = z And a79.Visible = z And a78.Visible = z And a77.Visible = z And a76.Visible = z And a75.Visible = z And a74.Visible = z And a73.Visible = z _
And a1.Visible = z And a2.Visible = z And a3.Visible = z And a4.Visible = z And a10.Visible = z And a11.Visible = z And a12.Visible = z And a13.Visible = z And a19.Visible = z And a20.Visible = z And a21.Visible = z And a22.Visible = z And a28.Visible = z And a29.Visible = z And a30.Visible = z And a31.Visible = z And a37.Visible = z And a39.Visible = z And a40.Visible = z And a46.Visible = z And a47.Visible = z And a48.Visible = z And a49.Visible = z And Form1.Tiempo > 0 Then


Form1.Timer1.Enabled = False
Form1.Timer2.Enabled = False
Form2.Label1.Caption = "Su puntuación es: " + Form1.Tiempo + " segundos"

Form2.Text1.Text = ""
Form2.entrada.Caption = 0

Form2.Show 1

Timer6.Enabled = True

'Call nuevo
End If
End Sub


Private Sub a1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a1, Button)
End Sub

Private Sub a2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a2, Button)
End Sub

Private Sub a3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a3, Button)
End Sub

Private Sub a4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a4, Button)
End Sub

Private Sub a5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a5, Button)
End Sub

Private Sub a6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a6, Button)
End Sub

Private Sub a7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a7, Button)
End Sub

Private Sub a8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a8, Button)
End Sub
Private Sub a9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a9, Button)
End Sub

Private Sub a10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a10, Button)
End Sub

Private Sub a11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a11, Button)
End Sub

Private Sub a12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a12, Button)
End Sub

Private Sub a13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a13, Button)
End Sub

Private Sub a14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a14, Button)
End Sub

Private Sub a15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a15, Button)
End Sub

Private Sub a16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a16, Button)
End Sub

Private Sub a17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a17, Button)
End Sub

Private Sub a18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a18, Button)
End Sub

Private Sub a19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a19, Button)
End Sub

Private Sub a20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a20, Button)
End Sub

Private Sub a21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a21, Button)
End Sub

Private Sub a22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a22, Button)
End Sub

Private Sub a23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a23, Button)
End Sub

Private Sub a24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a24, Button)
End Sub

Private Sub a25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a25, Button)
End Sub

Private Sub a26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a26, Button)
End Sub

Private Sub a27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a27, Button)
End Sub

Private Sub a28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a28, Button)
End Sub

Private Sub a29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a29, Button)
End Sub

Private Sub a30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a30, Button)
End Sub

Private Sub a31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a31, Button)
End Sub

Private Sub a32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a32, Button)
End Sub

Private Sub a33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a33, Button)
End Sub

Private Sub a34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a34, Button)
End Sub

Private Sub a35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a35, Button)
End Sub

Private Sub a36_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a36, Button)
End Sub

Private Sub a37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a37, Button)
End Sub

Private Sub a38_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a38, Button)
End Sub

Private Sub a39_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a39, Button)
End Sub

Private Sub a40_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a40, Button)
End Sub

Private Sub a41_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a41, Button)
End Sub

Private Sub a42_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a42, Button)
End Sub

Private Sub a43_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a43, Button)
End Sub

Private Sub a44_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a44, Button)
End Sub

Private Sub a45_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a45, Button)
End Sub

Private Sub a46_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a46, Button)
End Sub

Private Sub a47_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a47, Button)
End Sub

Private Sub a48_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a48, Button)
End Sub

Private Sub a49_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a49, Button)
End Sub

Private Sub a50_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a50, Button)
End Sub

Private Sub a51_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a51, Button)
End Sub

Private Sub a52_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a52, Button)
End Sub

Private Sub a53_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a53, Button)
End Sub

Private Sub a54_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a54, Button)
End Sub

Private Sub a55_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a55, Button)
End Sub

Private Sub a56_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a56, Button)
End Sub

Private Sub a57_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a57, Button)
End Sub

Private Sub a58_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a58, Button)
End Sub

Private Sub a59_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a59, Button)
End Sub

Private Sub a60_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a60, Button)
End Sub

Private Sub a61_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a61, Button)
End Sub

Private Sub a62_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a62, Button)
End Sub

Private Sub a63_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a63, Button)
End Sub

Private Sub a64_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a64, Button)
End Sub

Private Sub a65_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a65, Button)
End Sub

Private Sub a66_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a66, Button)
End Sub

Private Sub a67_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a67, Button)
End Sub

Private Sub a68_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a68, Button)
End Sub

Private Sub a69_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a69, Button)
End Sub

Private Sub a70_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a70, Button)
End Sub

Private Sub a71_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a71, Button)
End Sub

Private Sub a72_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a72, Button)
End Sub

Private Sub a73_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a73, Button)
End Sub

Private Sub a74_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a74, Button)
End Sub

Private Sub a75_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a75, Button)
End Sub

Private Sub a76_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a76, Button)
End Sub

Private Sub a77_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a77, Button)
End Sub

Private Sub a78_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a78, Button)
End Sub

Private Sub a79_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a79, Button)
End Sub

Private Sub a80_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a80, Button)
End Sub

Private Sub a81_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call bandera(a81, Button)
End Sub


Sub activado()
If cambioo = 0 Then
If Frame1.Top >= 118 Then
Frame1.Top = Frame1.Top - 1
Else
cambioo = 1
End If
End If


If cambioo = 1 Then
If Shape2.Height <= 3255 Then
Shape2.Height = Shape2.Height + 15
Else
cambioo = 0
End If
End If

If cambioo = 1 Then
If Frame1.Height <= 217 Then
Frame1.Height = Frame1.Height + 1
Else
Timer4.Enabled = False
End If
End If

End Sub

Sub desactivado()

If cambioo = 0 Then
If Frame1.Top <= 200 Then
Frame1.Top = Frame1.Top + 1
Else
cambioo = 1
End If
End If

End Sub

Private Sub Timer3_Timer()
a2.Left = a1.Left + 32
a3.Left = a2.Left + 32
a4.Left = a3.Left + 32
a5.Left = a4.Left + 32
a6.Left = a5.Left + 32
a7.Left = a6.Left + 32
a8.Left = a7.Left + 32
a9.Left = a8.Left + 32
'------------------

a2.Top = a1.Top
a3.Top = a1.Top
a4.Top = a2.Top
a5.Top = a3.Top
a6.Top = a4.Top
a7.Top = a5.Top
a8.Top = a6.Top
a9.Top = a7.Top

a11.Top = a10.Top
a12.Top = a10.Top
a13.Top = a10.Top
a14.Top = a10.Top
a15.Top = a10.Top
a16.Top = a10.Top
a17.Top = a10.Top
a18.Top = a10.Top

a20.Top = a19.Top
a21.Top = a19.Top
a22.Top = a19.Top
a23.Top = a19.Top
a24.Top = a19.Top
a25.Top = a19.Top
a26.Top = a19.Top
a27.Top = a19.Top

a29.Top = a28.Top
a30.Top = a28.Top
a31.Top = a28.Top
a32.Top = a28.Top
a33.Top = a28.Top
a34.Top = a28.Top
a35.Top = a28.Top
a36.Top = a28.Top


a38.Top = a37.Top
a39.Top = a37.Top
a40.Top = a37.Top
a41.Top = a37.Top
a42.Top = a37.Top
a43.Top = a37.Top
a44.Top = a37.Top
a45.Top = a37.Top

a47.Top = a46.Top
a48.Top = a46.Top
a49.Top = a46.Top
a50.Top = a46.Top
a51.Top = a46.Top
a52.Top = a46.Top
a53.Top = a46.Top
a54.Top = a46.Top

a56.Top = a55.Top
a57.Top = a55.Top
a58.Top = a55.Top
a59.Top = a55.Top
a60.Top = a55.Top
a61.Top = a55.Top
a62.Top = a55.Top
a63.Top = a55.Top


a65.Top = a64.Top
a66.Top = a64.Top
a67.Top = a64.Top
a68.Top = a64.Top
a69.Top = a64.Top
a70.Top = a64.Top
a71.Top = a64.Top
a72.Top = a64.Top


a74.Top = a73.Top
a75.Top = a73.Top
a76.Top = a73.Top
a77.Top = a73.Top
a78.Top = a73.Top
a79.Top = a73.Top
a80.Top = a73.Top
a81.Top = a73.Top

a10.Left = a1.Left
a19.Left = a1.Left
a28.Left = a1.Left
a37.Left = a1.Left
a46.Left = a1.Left
a55.Left = a1.Left
a64.Left = a1.Left
a73.Left = a1.Left


a11.Left = a2.Left
a20.Left = a2.Left
a29.Left = a2.Left
a38.Left = a2.Left
a47.Left = a2.Left
a56.Left = a2.Left
a65.Left = a2.Left
a74.Left = a2.Left


a12.Left = a3.Left
a21.Left = a3.Left
a30.Left = a3.Left
a39.Left = a3.Left
a48.Left = a3.Left
a57.Left = a3.Left
a66.Left = a3.Left
a75.Left = a3.Left


a13.Left = a4.Left
a22.Left = a4.Left
a31.Left = a4.Left
a40.Left = a4.Left
a49.Left = a4.Left
a58.Left = a4.Left
a67.Left = a4.Left
a76.Left = a4.Left

a14.Left = a5.Left
a23.Left = a5.Left
a32.Left = a5.Left
a41.Left = a5.Left
a50.Left = a5.Left
a59.Left = a5.Left
a68.Left = a5.Left
a77.Left = a5.Left


a15.Left = a6.Left
a24.Left = a6.Left
a33.Left = a6.Left
a42.Left = a6.Left
a51.Left = a6.Left
a60.Left = a6.Left
a69.Left = a6.Left
a78.Left = a6.Left


a16.Left = a7.Left
a25.Left = a7.Left
a34.Left = a7.Left
a43.Left = a7.Left
a52.Left = a7.Left
a61.Left = a7.Left
a70.Left = a7.Left
a79.Left = a7.Left


a17.Left = a8.Left
a26.Left = a8.Left
a35.Left = a8.Left
a44.Left = a8.Left
a53.Left = a8.Left
a62.Left = a8.Left
a71.Left = a8.Left
a80.Left = a8.Left


a18.Left = a9.Left
a27.Left = a9.Left
a36.Left = a9.Left
a45.Left = a9.Left
a54.Left = a9.Left
a63.Left = a9.Left
a72.Left = a9.Left
a81.Left = a9.Left

End Sub

Private Sub Timer4_Timer()
If list.Caption = 1 Then
Call desactivado
End If

If list.Caption = 0 Then
Call activado
End If

End Sub

Private Sub Form_Load()
Call nuevo
Frame1.Visible = True

bor = 0
Timer5.Enabled = True

Call truco

End Sub


Private Sub Image4_Click()
Call sonidoplay(App.Path & "\recursos\7.wav", sincronizado)

Timer4.Enabled = True
list.Caption = 1
Call activado
Form1.Frame1.Visible = False
End Sub

Private Sub Timer5_Timer()
anim = anim + 1
Call vis
Label4.Caption = anim
End Sub

Private Sub Timer6_Timer()
actl.Caption = sumo
sumo = sumo + 1
Call invis
End Sub

