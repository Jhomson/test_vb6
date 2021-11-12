VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm1 
   Caption         =   "CRUD #1"
   ClientHeight    =   8535
   ClientLeft      =   165
   ClientTop       =   210
   ClientWidth     =   17385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8338.793
   ScaleMode       =   0  'User
   ScaleWidth      =   17385
   Begin VB.ComboBox searchbox 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "CRUD.frx":0000
      Left            =   4800
      List            =   "CRUD.frx":0002
      TabIndex        =   24
      Text            =   "Seleccione opción"
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CommandButton leerbd_btn 
      Caption         =   "Listar Datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   22
      Top             =   4920
      Width           =   3375
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   2895
      Left            =   4680
      TabIndex        =   21
      Top             =   720
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txti 
      Height          =   495
      Left            =   1920
      TabIndex        =   20
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox txth 
      Height          =   525
      Left            =   1920
      TabIndex        =   19
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtg 
      Height          =   495
      Left            =   1920
      TabIndex        =   18
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox txtf 
      Height          =   495
      Left            =   1920
      TabIndex        =   17
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox txte 
      Height          =   525
      Left            =   1920
      TabIndex        =   16
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtd 
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   1920
      Width           =   2535
   End
   Begin VB.ComboBox optionbox 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "CRUD.frx":0004
      Left            =   10560
      List            =   "CRUD.frx":0006
      TabIndex        =   7
      Text            =   "Seleccione opción"
      Top             =   3840
      Width           =   3495
   End
   Begin VB.CommandButton runbtn 
      Caption         =   "Ejecutar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14280
      TabIndex        =   6
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox txtc 
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtb 
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txta 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "Desarrollado por: Jhomson M. Arcas R. como prueba para empresa GUX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   28
      Top             =   240
      Width           =   7815
   End
   Begin VB.Label Label13 
      Caption         =   $"CRUD.frx":0008
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   11760
      TabIndex        =   27
      Top             =   5880
      Width           =   5415
   End
   Begin VB.Label Label12 
      Caption         =   $"CRUD.frx":0157
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   5040
      TabIndex        =   26
      Top             =   5880
      Width           =   6375
   End
   Begin VB.Label Label11 
      Caption         =   $"CRUD.frx":0325
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   25
      Top             =   5880
      Width           =   4335
   End
   Begin VB.Label Label10 
      Caption         =   "Seleccione el tipo de busqueda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   23
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label9 
      Caption         =   "Tipo PAM"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Nro."
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Acción"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Codigo pres"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha previsto"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha alta"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label itemslistlabel 
      Caption         =   "Lista de Items"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha hosp"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre paciente"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nro. Rol"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub inicializar()
    
    optionbox.Clear
    optionbox.Text = "Seleccione"
    optionbox.AddItem "Registrar", 0
    optionbox.AddItem "Actualizar", 1
    optionbox.AddItem "Eliminar", 2
    
    searchbox.Clear
    searchbox.Text = "Seleccione"
    searchbox.AddItem "Nro. Rol", 0
    searchbox.AddItem "Nombre del Paciente", 1
    
End Sub

Private Sub Form_Load()
    Call inicializar
End Sub

'El DSN lo he creado en el grupo de sistema (inicio/panel de control/Herramientas administrativas/Orígenes de datos (ODBC))
'Referencias actuales (2021-11-10)
'https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/movefirst-movelast-movenext-and-moveprevious-methods-example-vb

'Función para leer datos desde postgresql y colocarlos en campos dentro del formulario, así como en datagrids.
Private Sub leerdatos(consulta As String)

    Dim cN As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim strCn As String
    Dim il_total As Integer, il_cont As Integer
    
    'Se establece la conexión a la BD
    Set cN = New ADODB.Connection
    'El dsn en el caso de windows se establece en los origenes de datos (ODBC)
    strCn = "dsn=PostgreSQL35W"
    cN.ConnectionString = strCn
    cN.Open
    'Se establece la sentencia deseada
    sl_txt = consulta
    'Se ejecuta la consulta a la BD
    rs.Open sl_txt, cN, adOpenStatic, adLockReadOnly
    'Se obtiene el total de registros
    il_total = rs.RecordCount
    
    'Se inserta la data dentro del datagrid en el formulario
    Set DataGrid2.DataSource = rs
    DataGrid2.Refresh
    
    'A modo de ejemplo, de esta forma se pueden leer los datos directamente
    'para colocarlos en cualquier elemento del formulario.
    'If il_total > 0 Then
    '    For il_cont = 1 To rs.RecordCount
    '        txta.Text = Trim(rs.Fields(0)) '--> Leyendo el Nro. Rol
    '        txtb.Text = Trim(rs.Fields(1)) '--> Leyendo el Nombre del Paciente
    '        txtc.Text = Trim(rs.Fields(2)) '--> Leyendo el Fecha Hosp.
    '        txtd.Text = Trim(rs.Fields(3)) '--> Leyendo el Fecha Alta
    '        txte.Text = Trim(rs.Fields(4)) '--> Leyendo el Fecha pres.
    '        txtf.Text = Trim(rs.Fields(5)) '--> Leyendo el Codigo
    '        txtg.Text = Trim(rs.Fields(6)) '--> Leyendo el Acción
    '        txth.Text = Trim(rs.Fields(7)) '--> Leyendo el Nro
    '        txti.Text = Trim(rs.Fields(8)) '--> Leyendo el Tipo PAM
    '        rs.MoveNext
    '    Next
    'Else
    '    MsgBox "No se encontro datos en la base de datos"
    'End If
    'rs.Close
End Sub

Private Sub leerbd_btn_Click()
    
    'Decidir según la opción seleccionada
    Dim opt As Integer
    Dim optselect As Integer
    Dim consulta As String
    Dim cada As String, cadb As String, cadc As String, cadd As String, cade As String
    Dim cadf As String, cadg As String, cadh As String, cadi As String
    
    cada = frm1.txta.Text
    cadb = frm1.txtb.Text
    cadc = frm1.txtc.Text
    cadd = frm1.txtd.Text
    cade = frm1.txte.Text
    cadf = frm1.txtf.Text
    cadg = frm1.txtg.Text
    cadh = frm1.txth.Text
    cadi = frm1.txti.Text
    
    If cada = "" Then
        cada = "?"
    End If
    If cadb = "" Then
        cadb = "?"
    End If
    If cadc = "" Then
        cadc = "?"
    End If
    If cadd = "" Then
        cadd = "?"
    End If
    If cade = "" Then
        cade = "?"
    End If
    If cadf = "" Then
        cadf = "?"
    End If
    If cadg = "" Then
        cadg = "?"
    End If
    If cadh = "" Then
        cadh = "?"
    End If
    If cadi = "" Then
        cadi = "?"
    End If
    
        
    opt = CInt(MsgBox("¿Desea proceder?", 1 + 512 + 32, "Caja de mensajes"))
    
    If opt = 1 Then
        'Decidir segun la opción seleccionada en la lista
        optselect = searchbox.ListIndex
        'MsgBox "Opción: " + CStr(optselect)
        
        'En caso de no haber rellenado alguna casilla para la busqueda, basta solo con agregar un caracter a cada valor para que la función like funcione perfectamente.
        'Se usa la clausula where junto con like para asegurar encontrar cualquier coincidencia segun lo que el Usuario indique al menos en un campo.
        'esto en conjunto con la instrucción lower para encontrar coincidencias sin importar mayusculas y minusculas.
        
        Select Case optselect
            
            Case Is < 0
                consulta = "select ""NROL"",""NOMBRE_PACIENTE"",to_char(""FECHA_HOSP"", 'YYYY-MM-DD'),to_char(""FECHA_ALTA"", 'YYYY-MM-DD'),to_char(""FECHA_PROVIS"", 'YYYY-MM-DD'),""CODIGO_PRES"",""ACCION"",""NRO"",""TIPO_PAM"" from testtable WHERE LOWER(CAST(""NROL"" AS text)) LIKE LOWER('" & cada & "%') OR LOWER(CAST(""NOMBRE_PACIENTE"" AS text)) LIKE LOWER('" & cadb & "%') OR LOWER(CAST(""FECHA_HOSP"" AS text)) LIKE LOWER('" & cadc & "%') OR LOWER(CAST(""FECHA_ALTA"" AS text)) LIKE LOWER('" & cadd & "%') OR LOWER(CAST(""FECHA_PROVIS"" AS text)) LIKE LOWER('" & cade & "%') OR LOWER(CAST(""CODIGO_PRES"" AS text)) LIKE LOWER('" & cadf & "%') OR LOWER(CAST(""ACCION"" AS text)) LIKE LOWER('" & cadg & "%') OR LOWER(CAST(""NRO"" AS text)) LIKE LOWER('" & cadh & "%') OR LOWER(CAST(""TIPO_PAM"" AS text)) LIKE LOWER('" & cadi & "%') ORDER BY ""NOMBRE_PACIENTE"" ASC LIMIT 100;" '-->Se limita para no saturar el sistema si se trata de muchos datos, lo ideal sería colocar una función para exportar todos los datos.
                leerdatos (consulta)
            Case 0 '-->Busqueda por numero de rol (NROL)
                consulta = "select ""NROL"",""NOMBRE_PACIENTE"",to_char(""FECHA_HOSP"", 'YYYY-MM-DD'),to_char(""FECHA_ALTA"", 'YYYY-MM-DD'),to_char(""FECHA_PROVIS"", 'YYYY-MM-DD'),""CODIGO_PRES"",""ACCION"",""NRO"",""TIPO_PAM"" from testtable WHERE ""NROL"" = " & frm1.txta.Text & " ORDER BY ""NOMBRE_PACIENTE"" ASC;"
                leerdatos (consulta)
            Case 1 '-->Busqueda por nombre del paciente (NOMBRE_PACIENTE)
                consulta = "select ""NROL"",""NOMBRE_PACIENTE"",to_char(""FECHA_HOSP"", 'YYYY-MM-DD'),to_char(""FECHA_ALTA"", 'YYYY-MM-DD'),to_char(""FECHA_PROVIS"", 'YYYY-MM-DD'),""CODIGO_PRES"",""ACCION"",""NRO"",""TIPO_PAM"" from testtable ORDER BY ""NOMBRE_PACIENTE"" ASC WHERE ""NOMBRE_PACIENTE"" = " & frm1.txtb.Text & ";"
                leerdatos (consulta)
        End Select
    
    Else
        MsgBox "Verifique e intente nuevamente"
    End If
    
End Sub

Private Function limpiar_campos()

frm1.txta.Text = ""
frm1.txtb.Text = ""
frm1.txtc.Text = ""
frm1.txtd.Text = ""
frm1.txte.Text = ""
frm1.txtf.Text = ""
frm1.txtg.Text = ""
frm1.txth.Text = ""
frm1.txti.Text = ""

End Function

Private Sub registro_datos(cada As String, cadb As String, cadc As String, cadd As String, cade As String, cadf As String, cadg As String, cadh As String, cadi As String)
    Dim rs As New ADODB.Recordset
    Dim strCn As String
    Dim cN As ADODB.Connection
    Dim sMsg As String

    On Error GoTo ErrHandler

    Set cN = New ADODB.Connection
    strCn = "dsn=PostgreSQL35W"
    cN.ConnectionString = strCn
    cN.Open
    
    'Se establece la sentencia deseada
    sl_txt = "INSERT INTO testtable VALUES (" & cada & ",'" & cadb & "','" & cadc & "','" & cadd & "','" & cade & "','" & cadf & "','" & cadg & "'," & cadh & ",'" & cadi & "');"
    'Se ejecuta la consulta a la BD
    rs.Open sl_txt, cN, adOpenStatic, adLockReadOnly
    
    MsgBox "Registro Exito!"
    leerdatos
    limpiar_campos
    
    Exit Sub

ErrHandler:
    sMsg = "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
    'Verificamos cualquier error producido
    MsgBox sMsg
    
End Sub

Private Sub actualizar_datos(cada As String, cadb As String, cadc As String, cadd As String, cade As String, cadf As String, cadg As String, cadh As String, cadi As String)
    Dim rs As New ADODB.Recordset
    Dim strCn As String
    Dim cN As ADODB.Connection
    Dim sMsg As String

    On Error GoTo ErrHandler

    Set cN = New ADODB.Connection
    strCn = "dsn=PostgreSQL35W"
    cN.ConnectionString = strCn
    cN.Open
    
    'Se establece la sentencia deseada
    sl_txt = "UPDATE testtable SET ""NOMBRE_PACIENTE"" = '" & cadb & "', ""FECHA_HOSP"" = '" & cadc & "', ""FECHA_ALTA"" = '" & cadd & "', ""FECHA_PROVIS"" = '" & cade & "', ""CODIGO_PRES"" = '" & cadf & "', ""ACCION"" = '" & cadg & "', ""NRO"" = " & cadh & ", ""TIPO_PAM"" = '" & cadi & "' WHERE ""NROL"" = " & cada & ";"
    'Se ejecuta la consulta a la BD
    rs.Open sl_txt, cN, adOpenStatic, adLockReadOnly
    
    MsgBox "Actualización Exitosa!"
    leerdatos
    limpiar_campos
    
    Exit Sub

ErrHandler:
    sMsg = "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
    'Verificamos cualquier error producido
    MsgBox sMsg
    
End Sub

Private Sub eliminar_datos(cada As String)
    Dim rs As New ADODB.Recordset
    Dim strCn As String
    Dim cN As ADODB.Connection
    Dim sMsg As String

    On Error GoTo ErrHandler

    Set cN = New ADODB.Connection
    strCn = "dsn=PostgreSQL35W"
    cN.ConnectionString = strCn
    cN.Open
    
    'Se establece la sentencia deseada
    sl_txt = "DELETE FROM testtable WHERE ""NROL"" = " & cada & ";"
    'Se ejecuta la consulta a la BD
    rs.Open sl_txt, cN, adOpenStatic, adLockReadOnly
    
    MsgBox "Eliminado con Exito!"
    leerdatos
    limpiar_campos
    
    Exit Sub

ErrHandler:
    sMsg = "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
    'Verificamos cualquier error producido
    MsgBox sMsg
    
End Sub

Private Sub DataGrid2_DblClick()
    Me.DataGrid2.Row = Me.DataGrid2.Row
    
    Me.DataGrid2.Col = 0
    frm1.txta = Me.DataGrid2.Text
    
    Me.DataGrid2.Col = 1
    frm1.txtb = Me.DataGrid2.Text
    
    Me.DataGrid2.Col = 2
    frm1.txtc = Me.DataGrid2.Text
    
    Me.DataGrid2.Col = 3
    frm1.txtd = Me.DataGrid2.Text
    
    Me.DataGrid2.Col = 4
    frm1.txte = Me.DataGrid2.Text
    
    Me.DataGrid2.Col = 5
    frm1.txtf = Me.DataGrid2.Text
    
    Me.DataGrid2.Col = 6
    frm1.txtg = Me.DataGrid2.Text
    
    Me.DataGrid2.Col = 7
    frm1.txth = Me.DataGrid2.Text
    
    Me.DataGrid2.Col = 8
    frm1.txti = Me.DataGrid2.Text
End Sub

'Se establece el listado de acciones que se ejecutaran al seleccionar cada una de las opciones disponibles

Private Sub runbtn_Click()

    'Decidir según la opción seleccionada
    Dim opt As Integer
    Dim optselect As Integer
    
    opt = CInt(MsgBox("¿Desea proceder?", 1 + 512 + 32, "Caja de mensajes"))
    
    If opt = 1 Then
        'Decidir segun la opción seleccionada en la lista
        optselect = optionbox.ListIndex
        'MsgBox "Opción: " + CStr(optselect)
        
        If optselect < 0 Then
            MsgBox "Debe seleccionar una opción valida!"
            Exit Sub
        End If
        
        If optselect = 0 Then ' --> Registro
            'Se va a registrar la data en la base de datos
            Call registro_datos(frm1.txta, frm1.txtb, frm1.txtc, frm1.txtd, frm1.txte, frm1.txtf, frm1.txtg, frm1.txth, frm1.txti)
        End If
        
        If optselect = 1 Then ' --> Actualización
            'Se va a registrar la data en la base de datos
            Call actualizar_datos(frm1.txta, frm1.txtb, frm1.txtc, frm1.txtd, frm1.txte, frm1.txtf, frm1.txtg, frm1.txth, frm1.txti)
        End If
        
        If optselect = 2 Then ' --> Actualización
            'Se va a registrar la data en la base de datos
            Call eliminar_datos(frm1.txta)
        End If
        
    Else
        MsgBox "Verifique e intente nuevamente"
    End If
    
End Sub
