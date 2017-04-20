VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
   ' Api AnimateWindow
    Private Declare Function AnimateWindow Lib "user32" ( _
        ByVal hwnd As Long, ByVal dwTime As Long, _
        ByVal dwFlags As Long) As Long
     
    ' Api SetWindowPos
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
        ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
     
    ' Constantes para SetWindowPos
    Const SWP_FRAMECHANGED = &H20
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    Const SWP_Flags = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE
     
    Private Const AW_ACTIVATE = &H20000 'Activa la ventana cuando finaliza el efecto
     
    Enum ConstAnimation
        AW_BLEND = &H80000    ' Efecto fade
        AW_HOR_NEGATIVE = &H2 ' Derecha a izquierda
        AW_HOR_POSITIVE = &O1 ' Izquierda a derecha
        AW_VER_POSITIVE = &H4 ' Arriba abajo
        AW_VER_NEGATIVE = &H8 ' Abajo hacia arriba
        AW_HIDE = &H10000
        AW_CENTER = &H10      ' Aparece desde el Centro
    End Enum
     
    Public Event Error(nErr As Long, Description As String, Source As String)
    Public Event InitializeAnimation()
    Public Event TerminateAnimation()
     
    Public Function Animate(Control As Object, MiliSecond As Long, Animacion As ConstAnimation) As Long
    On Error GoTo ErrorR
        RaiseEvent InitializeAnimation      ' LLamada a evento (Antes de Iniciar)
        Control.Visible = False             ' El Obj nesesita ser Invisible para poderla Cargar con una Animacion
        DoEvents
        Animate = AnimateWindow(Control.hwnd, MiliSecond, Animacion Or AW_ACTIVATE) ' Iniciar Animacion
        Control.Refresh                     ' Refrescar control
        RaiseEvent TerminateAnimation       ' LLamada a Evento
        Exit Function
     
ErrorR:
        ' Si ocurre un Error Iniciar evento
        RaiseEvent Error(Err.Number, Err.Description, Err.Source)
        Err.Clear
    End Function
     
    Public Function RefreshTotal(Frm As Form) As Long
    On Error Resume Next
    Dim ObjCont As Object
    For Each ObjCont In Frm.Controls
        ObjCont.Refresh
        Call SetWindowPos(ObjCont.hwnd, 0, 0, 0, 0, 0, SWP_Flags) ' Metodo para refrescar por hWnd
        DoEvents
    Next
    End Function

