VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frm2 
   Caption         =   "PokeAPI Visual Basic"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15825
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   15825
   Begin VB.TextBox endbox 
      Height          =   375
      Left            =   11880
      TabIndex        =   8
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox initbox 
      Height          =   375
      Left            =   9960
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5535
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   12000
      ExtentX         =   21167
      ExtentY         =   9763
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton nextbtn 
      Caption         =   "Siguiente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   12360
      TabIndex        =   3
      Top             =   960
      Width           =   3285
   End
   Begin VB.CommandButton backbtn 
      Caption         =   "Anterior"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   12360
      TabIndex        =   4
      Top             =   1800
      Width           =   3285
   End
   Begin VB.TextBox cantfield 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton getpokebtn 
      Caption         =   "Listar Pokemon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "------->"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   9
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Se muestran los items:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad de elementos a mostrar       (Deje en blanco para mostrar de 10 en 10)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable publica que puede ser usada desde otro formulario.
Public valor As String

Private Sub Form_Load()
    Me.Caption = valor
    WebBrowser1.Navigate "about:blank"
    frm2.endbox.Enabled = False
    frm2.initbox.Enabled = False
    frm2.backbtn.Enabled = False
    frm2.nextbtn.Enabled = False
End Sub

Private Sub getpokebtn_Click()
    
    'Se deben deshabilitar los botones de siguiente y anterior para evitar errores
    frm2.backbtn.Enabled = False
    frm2.nextbtn.Enabled = False
    'Aca se comienza a recibir toda la información de los Pokemon.
    get_pokeinfo
    
End Sub

Public Sub navegar_lista(cant_mostrada As Integer, texto As String, backornext As Integer)
    
    'Dependiendo del valor de la variable cantidad, se va a mostrar el numero de items deseado por pagina
    'Del item obj se debe hacer otro get para obtener asi la imagen de cada item
    
    Dim html As String
    Dim data_poke() As String
    Dim obj As Object
    Dim objimg As Object
    Dim imgurl As String
    Dim tempurl As String
    Dim count As Integer
    Dim initlist As Integer
    Dim endlist As Integer
    Dim puntero As Integer
    
    Set obj = JSON.parse(texto)
    count = CInt(obj.Item("count"))
    ReDim data_poke(count, count)
    Set httpURL_b = New WinHttp.WinHttpRequest
    
    If frm2.nextbtn.Enabled = False And frm2.backbtn.Enabled = False Then
        initlist = 1
        
        If frm2.cantfield.Text = "" Then
            endlist = cant_mostrada 'Si el Usuario no especifica cuantos elementos desea ver, se muestra la cantidad por defecto (10)
            frm2.endbox.Text = cant_mostrada 'Para evitar error de limite de elementos
            frm2.cantfield.Text = cant_mostrada 'Colocar valor por defecto para evitar el error al pasar de página
        Else
            endlist = CInt(frm2.cantfield.Text)
            If backornext = 0 Then
                frm2.endbox.Text = frm2.cantfield.Text
            End If
        End If

        frm2.nextbtn.Enabled = True
        frm2.backbtn.Enabled = True
        frm2.initbox.Text = 1
    Else
        'Esto una vez que se ha sacado la cuenta y se establece el nuevo valor.
        If frm2.cantfield.Text = "" Then
            endlist = cant_mostrada 'Si el Usuario no especifica cuantos elementos desea ver, se muestra la cantidad por defecto (10)
            frm2.endbox.Text = cant_mostrada 'Para evitar error de limite de elementos
            frm2.cantfield.Text = cant_mostrada 'Colocar valor por defecto para evitar el error al pasar de página
        Else
            endlist = CInt(frm2.cantfield.Text)
            If backornext = 0 Then
                frm2.endbox.Text = frm2.cantfield.Text
            End If
        End If
        
        initlist = CInt((frm2.initbox.Text))
        endlist = CInt((frm2.endbox.Text))
        'MsgBox "Inicio: " & initlist & " Final: " & endlist
    End If
    
    'Se debe deshabilitar el boton de regresar en caso de que se haya llegado al inicio de la lista
    If frm2.initbox.Text = "1" Then
        frm2.backbtn.Enabled = False
    Else
        frm2.backbtn.Enabled = True
        frm2.nextbtn.Enabled = True
    End If
    
    'Se debe construir un Arreglo que alamacene todos los elementos necesarios, en este caso solo el nombre y la imagen.
    puntero = 1 'Para comenzar a guardar desde el elemento numero 1.
    For rec = initlist To endlist
    
        'Se debe realizar otro get para obtener la información adicional del Pokemon en cuestión
        'En teoría este debería ser el método mas adecuado porque en caso de que se cambie la url de igual forma se
        'seguirá obteniendo la url de las imagenes, sin embargo, se genera un poco mas de lentitud por la cantidad
        'de peticiones al servidor, para hacer el proceso mas rápido, puede probar a comentar las lineas indicadas como "método lento" mas abajo
        'y descomentar las indicadas como "método rápido" para observar como aumenta la velocidad de respuesta
        'puesto que ya se sabe la url donde se guardan las imagenes de los items.
        
        'Metodo lento (Petición HTTP)
        tempurl = obj.Item("results").Item(rec).Item("url")
        'Metodo lento (Petición HTTP)
        httpURL_b.Open "GET", tempurl
        httpURL_b.Send
        imgurl = httpURL_b.ResponseText
        'Metodo lento (Petición HTTP)
        If imgurl = "[]" Then
            MsgBox ("No se obtuvo información")
            Exit Sub
        End If
    
        'En caso de recibir información
        'Metodo lento (Petición HTTP)
        Set objimg = JSON.parse(imgurl)
        
        'Metodo lento (Petición HTTP)
        data_poke(puntero, 1) = objimg.Item("sprites").Item("front_default")
        
        'Método rápido (URL directa)
        'data_poke(puntero, 1) = "https://raw.githubusercontent.com/PokeAPI/sprites/master/sprites/pokemon/" & rec & ".png"
        
        data_poke(puntero, 0) = obj.Item("results").Item(rec).Item("name")
        puntero = puntero + 1
        On Error GoTo final_list
    
    Next rec
    
    'MsgBox "Pokemon: " & data_poke(1, 0) & " Imagen: " & data_poke(1, 1)
    
    'Codigo html de la interfaz
html_make:

    html = "<html>" & _
            "<head>" & _
                "<style type='text/css'>" & _
                    "#titlediv{color: white; background-color: black; font-weight: bold; font-size: 20px;} #container_global{display:flex; align-items: center; justify-content: center;} #container_global div{float: left; margin: 10px; background-color: silver;} #container_global div img{width: 100%; height: 100px;}" & _
                "</style>" & _
            "</head>" & _
            "<body>" & _
                "<div id='titlediv'><p align = 'center' style='margin:20;padding:20;'>Kanto Pokemon</p></div>" & _
                "<div id='container_global' style='width:755px; height:600px; border: 1px solid transparent;'>"
                
    For reg = 1 To cant_mostrada 'Solo se muestran la cantidad de elementos que el Usuario ha especificado.
    
    'Se establece una consición que impida imprimir valores no deseados.
    If data_poke(reg, 0) = "" Then
        'No imprimir, saldría una div vacio.
    Else
    
        html = html & "<div style='width:165px; height:200px;'>" & _
                        "<img src=""" & data_poke(reg, 1) & """>" & _
                        "<p align = 'center'><strong>" & data_poke(reg, 0) & "</strong></p>" & _
                        "</div>"
    End If
    
    Next reg
                
    html = html & "</div>" & _
                "</body>" & _
                "<script>alert('hola');<script>" & _
            "</html>"
    
    'Se debe construir una lista de items cumpliendo con la cantidad de items que haya establecido el Usuario
    'en el campo "cantfield"
    On Error GoTo refresh_forced
        
        WebBrowser1.Document.Write html
        WebBrowser1.Refresh
        Exit Sub
        
refresh_forced:
        
        frm2.nextbtn.Enabled = False
        WebBrowser1.Document.Write html
        WebBrowser1.Refresh
        Exit Sub
        
final_list:
        frm2.nextbtn.Enabled = False
        GoTo html_make
        
ver_error:
    MsgBox Err.Description
    frm2.nextbtn.Enabled = False

End Sub

Private Function textoconsulta(cadena As String) As String
    Set httpURL = New WinHttp.WinHttpRequest
    httpURL.Open "GET", cadena
    httpURL.Send
    textoconsulta = httpURL.ResponseText
End Function

Private Function get_pokeinfo()

    Dim obj As Object
    Dim texto As String
    Dim sinputjson As String
    Dim nombre As String
    Dim url As String
    Dim count As String
    
    'Set httpURL = New WinHttp.WinHttpRequest

    'cadena = "https://pokeapi.co/api/v2/pokemon?limit=151"
    'httpURL.Open "GET", cadena
    'httpURL.Send
    texto = textoconsulta("https://pokeapi.co/api/v2/pokemon?limit=151")
   
    If texto = "[]" Then
        MsgBox ("No se obtuvo información")
        Exit Function
    End If
    
    If IsNumeric(frm2.cantfield.Text) Then
        Call navegar_lista(CInt(frm2.cantfield.Text), texto, 0)
    Else
        Call navegar_lista(10, texto, 0) 'Valor por defecto
    End If
    
End Function

Private Sub nextbtn_Click()
    Dim texto As String
    
    'Se debe calcular el avance en la lista
    frm2.initbox.Text = CStr(CInt(frm2.cantfield.Text) + CInt(frm2.initbox.Text))
    frm2.endbox.Text = CStr(CInt(frm2.cantfield.Text) + CInt(frm2.endbox.Text))
    texto = textoconsulta("https://pokeapi.co/api/v2/pokemon?limit=151")
    Call navegar_lista(CInt(frm2.cantfield.Text), texto, 1)
End Sub

Private Sub backbtn_Click()
    Dim texto As String
    
    'Se debe calcular el retroceso en la lista
    frm2.endbox.Text = CStr(CInt(frm2.endbox.Text) - CInt(frm2.cantfield.Text))
    frm2.initbox.Text = CStr(CInt(frm2.initbox.Text) - CInt(frm2.cantfield.Text))
    texto = textoconsulta("https://pokeapi.co/api/v2/pokemon?limit=151")
    Call navegar_lista(CInt(frm2.cantfield.Text), texto, 1)
End Sub
