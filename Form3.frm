VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambiar apariencia"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7965
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   531
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   6120
      Width           =   7695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4215
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccionar color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4440
      TabIndex        =   19
      Top             =   4080
      Width           =   3375
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   20
         TabIndex        =   40
         Top             =   1440
         Width           =   3255
         Begin VB.OptionButton fc7 
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   240
            TabIndex        =   43
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton fc8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Segoe Script"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Left            =   1320
            TabIndex        =   42
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton fc9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Segoe Script"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   375
            Left            =   2280
            TabIndex        =   41
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   20
         TabIndex        =   36
         Top             =   930
         Width           =   3255
         Begin VB.OptionButton fc4 
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00008000&
            Height          =   360
            Left            =   240
            TabIndex        =   39
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton fc5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Segoe Print"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   375
            Left            =   1320
            TabIndex        =   38
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton fc6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Segoe Print"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   375
            Left            =   2280
            TabIndex        =   37
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   20
         TabIndex        =   32
         Top             =   360
         Width           =   3255
         Begin VB.OptionButton fc2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Segoe Print"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   375
            Left            =   1320
            TabIndex        =   35
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton fc3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Segoe Script"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   375
            Left            =   2280
            TabIndex        =   34
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton fc1 
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   0
            Width           =   735
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccionar fuente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1935
      Left            =   120
      TabIndex        =   18
      Top             =   4080
      Width           =   3975
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   3735
         Begin VB.OptionButton f7 
            BackColor       =   &H00FFFFFF&
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
            Height          =   360
            Left            =   240
            TabIndex        =   31
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton f8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   30
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton f9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Segoe Script"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            TabIndex        =   29
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   930
         Width           =   3735
         Begin VB.OptionButton f4 
            BackColor       =   &H00FFFFFF&
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
            Height          =   360
            Left            =   240
            TabIndex        =   27
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton f5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   26
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton f6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Segoe Script"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            TabIndex        =   25
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   3735
         Begin VB.OptionButton f1 
            BackColor       =   &H00FFFFFF&
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
            Height          =   360
            Left            =   240
            TabIndex        =   23
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton f2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   22
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton f3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Segoe Script"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            TabIndex        =   21
            Top             =   0
            Width           =   615
         End
      End
   End
   Begin VB.CommandButton a2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2"
      Height          =   480
      Left            =   6840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   480
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vista Previa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4440
      TabIndex        =   12
      Top             =   360
      Width           =   3375
      Begin VB.CommandButton a1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
         Height          =   480
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   480
      End
      Begin VB.Image bomb 
         Height          =   480
         Index           =   0
         Left            =   960
         Picture         =   "Form3.frx":0000
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image bomb 
         Height          =   480
         Index           =   1
         Left            =   1920
         Picture         =   "Form3.frx":05BC
         Top             =   2400
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   960
         X2              =   960
         Y1              =   480
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   1
         X1              =   1440
         X2              =   1440
         Y1              =   480
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   2
         X1              =   1920
         X2              =   1920
         Y1              =   480
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   3
         X1              =   2400
         X2              =   2400
         Y1              =   480
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   4
         X1              =   480
         X2              =   2880
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   5
         X1              =   480
         X2              =   2880
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   6
         X1              =   480
         X2              =   2880
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   7
         X1              =   480
         X2              =   2880
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   495
         Index           =   0
         Left            =   1110
         TabIndex        =   16
         Top             =   1430
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   465
         Index           =   1
         Left            =   1590
         TabIndex        =   15
         Top             =   1440
         Width           =   210
      End
      Begin VB.Label Label1 
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
         Left            =   2085
         TabIndex        =   14
         Top             =   1500
         Width           =   180
      End
      Begin VB.Shape Shape1 
         Height          =   2415
         Left            =   480
         Top             =   480
         Width           =   2415
      End
      Begin VB.Image fondo 
         Height          =   2415
         Left            =   480
         Picture         =   "Form3.frx":0B78
         Stretch         =   -1  'True
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccionar fondo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   3975
      Begin VB.OptionButton Option9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Option3"
         Height          =   195
         Left            =   2760
         TabIndex        =   11
         Top             =   600
         Width           =   230
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
      Begin VB.Image fondo1 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   600
         Stretch         =   -1  'True
         Top             =   480
         Width           =   480
      End
      Begin VB.Image fondo1 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   2
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   480
         Width           =   480
      End
      Begin VB.Image fondo1 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   1
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccionar mina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   3975
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   255
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Option3"
         Height          =   195
         Left            =   2760
         TabIndex        =   5
         Top             =   600
         Width           =   255
      End
      Begin VB.Image mina 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   600
         Stretch         =   -1  'True
         Top             =   480
         Width           =   480
      End
      Begin VB.Image mina 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   2
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   480
         Width           =   480
      End
      Begin VB.Image mina 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   1
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   435
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccionar tablero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Option3"
         Height          =   195
         Left            =   2760
         TabIndex        =   3
         Top             =   600
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   600
         Width           =   255
      End
      Begin VB.Image tabl 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   600
         Stretch         =   -1  'True
         Top             =   480
         Width           =   480
      End
      Begin VB.Image tabl 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   2
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   480
         Width           =   480
      End
      Begin VB.Image tabl 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   1
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.Image tr5 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   -15
      Picture         =   "Form3.frx":7A09
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   480
   End
   Begin VB.Image tr4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   0
      Picture         =   "Form3.frx":7F2A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
   Begin VB.Image tr3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   7305
      Picture         =   "Form3.frx":844B
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image tr2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   4455
      Picture         =   "Form3.frx":896C
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   480
   End
   Begin VB.Image tr1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   7485
      Picture         =   "Form3.frx":8E8D
      Stretch         =   -1  'True
      Top             =   -15
      Width           =   480
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pee As Integer
Dim pek As Integer
Dim trucao As Integer

Sub truco(contro As Image)
Dim trr As Control
Set trr = contro

If trucao >= 0 And trucao < 5 Then
'MsgBox trucao
trucao = trucao + 1
trr.Visible = False

If trucao = 5 Then
MsgBox "Ahora puede seleccionar una imagen y colocarla como tablero", vbInformation, "Truco activado"

With CommonDialog1
.Filter = "*.JPG | *.jpg"
.ShowOpen
End With

If CommonDialog1.FileName <> "" Then
Form1.tabl(1).Picture = LoadPicture(CommonDialog1.FileName)
Form3.tabl(1).Picture = Form1.tabl(1).Picture
Option2_Click
Form1.cam = 3
Form1.cam = 1

trucao = 0
tr1.Visible = True
tr2.Visible = True
tr3.Visible = True
tr4.Visible = True
tr5.Visible = True
Else
MsgBox "No seleccionó ninguna imágen." & vbCrLf & "Ahora deberá volver a activar truco", vbInformation, "Buscaminas"
trucao = 0
tr1.Visible = True
tr2.Visible = True
tr3.Visible = True
tr4.Visible = True
tr5.Visible = True

End If
End If

Exit Sub
End If

End Sub

Private Sub Command1_Click()
Form3.Visible = False
End Sub

Private Sub Form_Load()
tr1.BorderStyle = 0
tr2.BorderStyle = 0
tr3.BorderStyle = 0
tr4.BorderStyle = 0
tr5.BorderStyle = 0


Form1.Timer6.Enabled = True
Form3.tabl(0).Picture = Form1.tabl(0).Picture
Form3.tabl(1).Picture = Form1.tabl(1).Picture
Form3.tabl(2).Picture = Form1.tabl(2).Picture

Form3.mina(0).Picture = Form1.mina(0).Picture
Form3.mina(1).Picture = Form1.mina(1).Picture
Form3.mina(2).Picture = Form1.mina(2).Picture

Form3.fondo1(0).Picture = Form1.fondo1(0).Picture
Form3.fondo1(1).Picture = Form1.fondo1(1).Picture
Form3.fondo1(2).Picture = Form1.fondo1(2).Picture

Form3.fondo1(0).BorderStyle = Form1.fondo1(0).BorderStyle
Form3.fondo1(1).BorderStyle = Form1.fondo1(1).BorderStyle
Form3.fondo1(2).BorderStyle = Form1.fondo1(2).BorderStyle


If Form1.cam = 0 Then Option1.Value = True
If Form1.cam = 1 Then Option2.Value = True
If Form1.cam = 2 Then Option3.Value = True

If Form1.nbom.Caption = 0 Then Option4.Value = True
If Form1.nbom.Caption = 1 Then Option5.Value = True
If Form1.nbom.Caption = 2 Then Option6.Value = True

If Form1.fond = 0 Then Option7.Value = True
If Form1.fond = 1 Then Option8.Value = True
If Form1.fond = 2 Then Option9.Value = True

'-------------------
If Form1.Label1(0).Font = Form3.f1.Font Then Form3.f1.Value = True
If Form1.Label1(0).Font = Form3.f2.Font Then Form3.f2.Value = True
If Form1.Label1(0).Font = Form3.f3.Font Then Form3.f3.Value = True

If Form1.Label2(0).Font = Form3.f4.Font Then Form3.f4.Value = True
If Form1.Label2(0).Font = Form3.f5.Font Then Form3.f5.Value = True
If Form1.Label2(0).Font = Form3.f6.Font Then Form3.f6.Value = True

If Form1.Label3(0).Font = Form3.f7.Font Then Form3.f7.Value = True
If Form1.Label3(0).Font = Form3.f8.Font Then Form3.f8.Value = True
If Form1.Label3(0).Font = Form3.f9.Font Then Form3.f9.Value = True

'------------------
If Form1.Label1(0).ForeColor = Form3.fc1.ForeColor Then Form3.fc1.Value = True
If Form1.Label1(0).ForeColor = Form3.fc2.ForeColor Then Form3.fc2.Value = True
If Form1.Label1(0).ForeColor = Form3.fc3.ForeColor Then Form3.fc3.Value = True

If Form1.Label2(0).ForeColor = Form3.fc4.ForeColor Then Form3.fc4.Value = True
If Form1.Label2(0).ForeColor = Form3.fc5.ForeColor Then Form3.fc5.Value = True
If Form1.Label2(0).ForeColor = Form3.fc6.ForeColor Then Form3.fc6.Value = True

If Form1.Label3(0).ForeColor = Form3.fc7.ForeColor Then Form3.fc7.Value = True
If Form1.Label3(0).ForeColor = Form3.fc8.ForeColor Then Form3.fc8.Value = True
If Form1.Label3(0).ForeColor = Form3.fc9.ForeColor Then Form3.fc9.Value = True


End Sub

Private Sub Form_Unload(Cancel As Integer)
trucao = 0
Call sonidoplay(App.Path & "\recursos\7.wav", sincronizado)
End Sub

Private Sub Option1_Click()
Option1.Value = True
Form1.cam = 0
a1.Caption = ""
a1.Picture = tabl(0).Picture
a2.Caption = ""
a2.Picture = tabl(0).Picture

End Sub

Private Sub Option2_Click()
Option2.Value = True
Form1.cam = 1
a1.Caption = ""
a1.Picture = tabl(1).Picture
a2.Caption = ""
a2.Picture = tabl(1).Picture


End Sub

Private Sub Option3_Click()
Option3.Value = True
Form1.cam = 2
a1.Caption = ""
a1.Picture = tabl(2).Picture
a2.Caption = ""
a2.Picture = tabl(2).Picture
End Sub

Private Sub Option4_Click()
Option4.Value = True
Form1.nbom = 0
bomb(0).Picture = mina(0).Picture
bomb(1).Picture = mina(0).Picture

End Sub

Private Sub Option5_Click()
Option5.Value = True
Form1.nbom = 1
bomb(0).Picture = mina(1).Picture
bomb(1).Picture = mina(1).Picture
End Sub

Private Sub Option6_Click()
Option6.Value = True
Form1.nbom = 2
bomb(0).Picture = mina(2).Picture
bomb(1).Picture = mina(2).Picture
End Sub

Private Sub Option7_Click()
Option7.Value = True
Form1.fond = 0
fondo.Visible = True
fondo.Picture = fondo1(0).Picture
For pek = 0 To 7
Line1(pek).BorderColor = &HC0C0C0
Next pek
End Sub

Private Sub Option8_Click()
Option8.Value = True
Form1.fond = 1
fondo.Visible = True
fondo.Picture = fondo1(1).Picture
For pek = 0 To 7
Line1(pek).BorderColor = &HC0C0C0
Next pek
End Sub

Private Sub Option9_Click()
Dim pek
Option9.Value = True
Form1.fond = 2
fondo.Visible = True
fondo.Picture = fondo1(2).Picture

For pek = 0 To 7
Line1(pek).BorderColor = &HFFFFFF
Next pek

End Sub

Private Sub f1_Click()
f1.Value = True
Label1(0).Font = f1.Font
Label1(0).Top = 1430 + 70

For pee = 0 To 20
Form1.Label1(pee).Font = Form3.f1.Font
Next

Form1.Label1(0).Top = 14
Form1.Label1(1).Top = 14

Form1.Label1(2).Top = 46

Form1.Label1(3).Top = 78

Form1.Label1(4).Top = 110
Form1.Label1(5).Top = 110

Form1.Label1(6).Top = 142

Form1.Label1(7).Top = 172
Form1.Label1(8).Top = 206
Form1.Label1(9).Top = 206
Form1.Label1(10).Top = 206
Form1.Label1(11).Top = 206
Form1.Label1(12).Top = 206
Form1.Label1(13).Top = 238
Form1.Label1(14).Top = 238
Form1.Label1(15).Top = 238
Form1.Label1(16).Top = 270
Form1.Label1(17).Top = 270
Form1.Label1(18).Top = 270
Form1.Label1(19).Top = 270
Form1.Label1(20).Top = 270


End Sub

Private Sub f2_Click()

f2.Value = True
Label1(0).Font = f2.Font
Label1(0).Top = 1430 + 80

For pee = 0 To 20
Form1.Label1(pee).Font = Form3.f2.Font
Next

Form1.Label1(0).Top = 14 '- 4
Form1.Label1(1).Top = 14 '- 4
Form1.Label1(2).Top = 46 '- 4
Form1.Label1(3).Top = 78 '- 4
Form1.Label1(4).Top = 110 ' - 4
Form1.Label1(5).Top = 110 '- 4
Form1.Label1(6).Top = 142 '- 4
Form1.Label1(7).Top = 172 '- 4
Form1.Label1(8).Top = 206 '- 4
Form1.Label1(9).Top = 206 '- 4
Form1.Label1(10).Top = 206 '- 4
Form1.Label1(11).Top = 206 '- 4
Form1.Label1(12).Top = 206 '- 4
Form1.Label1(13).Top = 238 '- 4
Form1.Label1(14).Top = 238 '- 4
Form1.Label1(15).Top = 238 '- 4
Form1.Label1(16).Top = 270 '- 4
Form1.Label1(17).Top = 270 '- 4
Form1.Label1(18).Top = 270 '- 4
Form1.Label1(19).Top = 270 '- 4
Form1.Label1(20).Top = 270 '- 4


End Sub

Private Sub f3_Click()
f3.Value = True
Label1(0).Font = f3.Font
Label1(0).Top = 1430 + 10

For pee = 0 To 20
Form1.Label1(pee).Font = Form3.f3.Font
Next

Form1.Label1(0).Top = 14 - 3
Form1.Label1(1).Top = 14 - 3
Form1.Label1(2).Top = 46 - 3
Form1.Label1(3).Top = 78 - 3
Form1.Label1(4).Top = 110 - 3
Form1.Label1(5).Top = 110 - 3
Form1.Label1(6).Top = 142 - 3
Form1.Label1(7).Top = 172 - 3
Form1.Label1(8).Top = 206 - 3
Form1.Label1(9).Top = 206 - 3
Form1.Label1(10).Top = 206 - 3
Form1.Label1(11).Top = 206 - 3
Form1.Label1(12).Top = 206 - 3
Form1.Label1(13).Top = 238 - 3
Form1.Label1(14).Top = 238 - 3
Form1.Label1(15).Top = 238 - 3
Form1.Label1(16).Top = 270 - 3
Form1.Label1(17).Top = 270 - 3
Form1.Label1(18).Top = 270 - 3
Form1.Label1(19).Top = 270 - 3
Form1.Label1(20).Top = 270 - 3



End Sub

Private Sub f4_Click()
f4.Value = True
Label1(1).Font = f4.Font
Label1(1).Top = 1430 + 70

For pee = 0 To 10
Form1.Label2(pee).Font = Form3.f4.Font
Next


Form1.Label2(0).Top = 14
Form1.Label2(1).Top = 14
Form1.Label2(2).Top = 14

Form1.Label2(3).Top = 46

Form1.Label2(4).Top = 78
Form1.Label2(5).Top = 78
Form1.Label2(6).Top = 78

Form1.Label2(7).Top = 110
Form1.Label2(8).Top = 110

Form1.Label2(9).Top = 172

Form1.Label2(10).Top = 206
End Sub

Private Sub f5_Click()
f5.Value = True
Label1(1).Font = f5.Font
Label1(1).Top = 1430 + 80

For pee = 0 To 10
Form1.Label2(pee).Font = Form3.f5.Font
Next


Form1.Label2(0).Top = 14 '- 4
Form1.Label2(1).Top = 14 '- 4
Form1.Label2(2).Top = 14 '- 4
Form1.Label2(3).Top = 46 '- 4
Form1.Label2(4).Top = 78 '- 4
Form1.Label2(5).Top = 78 '- 4
Form1.Label2(6).Top = 78 '- 4
Form1.Label2(7).Top = 110 '- 4
Form1.Label2(8).Top = 110 '- 4
Form1.Label2(9).Top = 172 '- 4
Form1.Label2(10).Top = 206 '- 4


End Sub

Private Sub f6_Click()
f6.Value = True
Label1(1).Font = f6.Font
Label1(1).Top = 1430 + 10

For pee = 0 To 10
Form1.Label2(pee).Font = Form3.f6.Font
Next

Form1.Label2(0).Top = 14 - 3
Form1.Label2(1).Top = 14 - 3
Form1.Label2(2).Top = 14 - 3
Form1.Label2(3).Top = 46 - 3
Form1.Label2(4).Top = 78 - 3
Form1.Label2(5).Top = 78 - 3
Form1.Label2(6).Top = 78 - 3
Form1.Label2(7).Top = 110 - 3
Form1.Label2(8).Top = 110 - 3
Form1.Label2(9).Top = 172 - 3
Form1.Label2(10).Top = 206 - 3



End Sub

Private Sub f7_Click()



f7.Value = True
Label1(2).Font = f7.Font
Label1(2).Top = 1430 + 70

For pee = 0 To 4
Form1.Label3(pee).Font = Form3.f7.Font
Next

Form1.Label3(0).Top = 78
Form1.Label3(1).Top = 142
Form1.Label3(2).Top = 142
Form1.Label3(3).Top = 172
Form1.Label3(4).Top = 206


End Sub

Private Sub f8_Click()
f8.Value = True
Label1(2).Font = f8.Font
Label1(2).Top = 1430 + 80

For pee = 0 To 4
Form1.Label3(pee).Font = Form3.f8.Font
Next

Form1.Label3(0).Top = 78 '- 4
Form1.Label3(1).Top = 142 '- 4
Form1.Label3(2).Top = 142 '- 4
Form1.Label3(3).Top = 172 '- 4
Form1.Label3(4).Top = 206 '- 4


End Sub

Private Sub f9_Click()
f9.Value = True
Label1(2).Font = f9.Font
Label1(2).Top = 1430 + 10

For pee = 0 To 4
Form1.Label3(pee).Font = Form3.f9.Font
Next

Form1.Label3(0).Top = 78 - 3
Form1.Label3(1).Top = 142 - 3
Form1.Label3(2).Top = 142 - 3
Form1.Label3(3).Top = 172 - 3
Form1.Label3(4).Top = 206 - 3

End Sub

Private Sub fc1_Click()
fc1.Value = True
Label1(0).ForeColor = fc1.ForeColor

For pee = 0 To 20
Form1.Label1(pee).ForeColor = Form3.fc1.ForeColor
Next pee


End Sub

Private Sub fc2_Click()
fc2.Value = True
Label1(0).ForeColor = fc2.ForeColor

For pee = 0 To 20
Form1.Label1(pee).ForeColor = Form3.fc2.ForeColor
Next pee


End Sub

Private Sub fc3_Click()
fc3.Value = True
Label1(0).ForeColor = fc3.ForeColor

For pee = 0 To 20
Form1.Label1(pee).ForeColor = Form3.fc3.ForeColor
Next pee


End Sub

Private Sub fc4_Click()
fc4.Value = True
Label1(1).ForeColor = fc4.ForeColor

For pee = 0 To 10
Form1.Label2(pee).ForeColor = Form3.fc4.ForeColor
Next pee


End Sub

Private Sub fc5_Click()
fc5.Value = True
Label1(1).ForeColor = fc5.ForeColor

For pee = 0 To 10
Form1.Label2(pee).ForeColor = Form3.fc5.ForeColor
Next pee


End Sub

Private Sub fc6_Click()
fc6.Value = True
Label1(1).ForeColor = fc6.ForeColor

For pee = 0 To 10
Form1.Label2(pee).ForeColor = Form3.fc6.ForeColor
Next pee


End Sub

Private Sub fc7_Click()
fc7.Value = True
Label1(2).ForeColor = fc7.ForeColor

For pee = 0 To 4
Form1.Label3(pee).ForeColor = Form3.fc7.ForeColor
Next pee


End Sub

Private Sub fc8_Click()
fc8.Value = True
Label1(2).ForeColor = fc8.ForeColor

For pee = 0 To 4
Form1.Label3(pee).ForeColor = Form3.fc8.ForeColor
Next pee

End Sub

Private Sub fc9_Click()
fc9.Value = True
Label1(2).ForeColor = fc9.ForeColor

For pee = 0 To 4
Form1.Label3(pee).ForeColor = Form3.fc9.ForeColor
Next pee

End Sub



Private Sub a1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim a As String
a = Form1.cam
If Button = 2 Then
Call sonidoplay(App.Path & "\recursos\3.wav", sincronizado)

    If Form3.a1.Picture <> Form1.ban(a).Picture Then

            
                With Form3.a1
                    .Picture = Form1.ban(a).Picture
                    .Caption = ""
                    .ToolTipText = "Mina marcada"
                End With
     Else
    
    a1.Picture = tabl(a).Picture
    End If

End If

End Sub

Private Sub a2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim a As String
a = Form1.cam
If Button = 2 Then
Call sonidoplay(App.Path & "\recursos\3.wav", sincronizado)

    If Form3.a2.Picture <> Form1.ban(a).Picture Then

            
                With Form3.a2
                    .Picture = Form1.ban(a).Picture
                    .Caption = ""
                End With
     Else
    
    Form3.a2.Picture = tabl(a).Picture
    End If

End If

End Sub


Private Sub tr1_Click()
Call truco(tr1)
End Sub

Private Sub tr2_Click()
Call truco(tr2)
End Sub

Private Sub tr3_Click()
Call truco(tr3)
End Sub
Private Sub tr4_Click()
Call truco(tr4)
End Sub
Private Sub tr5_Click()
Call truco(tr5)
End Sub

