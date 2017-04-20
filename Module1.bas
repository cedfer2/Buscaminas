Attribute VB_Name = "Module1"
Declare Function sonidoplay Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal nombre_sonido As String, ByVal modo_de_reproduccion As Long) As Long


Public Const espera = &H0 'Reproduce un archivo de sonido, pero no devuelve el control al programa hasta que el archivo de sonido ha terminado de reproducir
Public Const sincronizado = &H1 ' reproduce en cualquier momento sin interrumpir programa
Public Const no_alarma = &H2 'reproducir archivo de audio y sino lo encuentra no sonar alarma
Public Const bucle = &H8 'reproducir con bucle
Public Const no_detener = &H10 'Asegura que si un archivo de sonido ya se está reproduciendo, el archivo de sonido no se interrumpe

Dim record As Recordset
Dim conex As Connection

Sub Cargar_ListView( _
        ListView As ListView, _
        sql As String, _
        PathBd As String)

    Dim Campo As Integer

    On Error GoTo ErrSub

    'Variable para los SubItem del LV
    Dim Item As ListItem
    Dim i As Long

    'Nuevo objeto Connection y objeto Recordset o contenedor de registros
    Set conex = New Connection
    Set record = New Recordset

    'Abre la base de datos
    conex.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PathBd & ";Persist Security Info=False")

    'Llena el Recordset
    record.Open sql, conex, adOpenDynamic, adLockOptimistic
    
    With ListView
        'Vista de reporte
        .View = lvwReport
        ' Elimina los item y los encabezado de columna
        .ListItems.Clear
        .ColumnHeaders.Clear
    End With

    Form5.MousePointer = 11

    'Agrega los nombres campo junto con los encabezados de columna para el ListView
    For Campo = 0 To record.Fields.Count - 1
        ListView.ColumnHeaders.Add , , record.Fields(Campo).Name
    Next

    ' Recorre todos los registros del Recordset
    While Not record.EOF
        'Agrega el Item
        Set Item = ListView.ListItems.Add(, , record.Fields(0))
        i = 1

        'Agrega los SubItem al ListView mediante la variable ITEM
        For Campo = 1 To record.Fields.Count - 1
    
            'si el dato no es de tipo Null lo agrega
            If Not IsNull(record.Fields(Campo)) Then
                Item.SubItems(i) = record.Fields(Campo)
            End If
            i = i + 1
        Next
    
    'Siguiente registro
    record.MoveNext
    Wend

    Form5.MousePointer = 0

Exit Sub
'Error
ErrSub:

    'MsgBox Err.Description, vbCritical, "Error"
    Form5.MousePointer = 0
End Sub





Sub cargar_list(lis As ListView)
sql = "SELECT * FROM puntuaciones order by tiempo asc"
midir = App.Path & "\recursos\db.mdb"

Call Cargar_ListView(lis, Trim$(sql), Trim$(midir))
End Sub


 Sub Ejecutar_Comando(lis As ListView, sql As String)
midir = App.Path & "\recursos\db.mdb"
Call Cargar_ListView(lis, Trim$(sql), Trim$(midir))

'trim quita espacio especificos ejemplo " hola " y con trim activada "hola"
'LTrim quita espacios finales ejemplo " hola " y con ltrim " hola"
 'rtrim quita espacions iniciales ejemplo " hola " y con rtrim "hola "
 End Sub
