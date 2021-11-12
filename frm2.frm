VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frm2 
   Caption         =   "PokeAPI Visual Basic"
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9540
   ScaleWidth      =   14265
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6615
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   13695
      ExtentX         =   24156
      ExtentY         =   11668
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
   Begin VB.CommandButton Command2 
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
      Height          =   700
      Left            =   9960
      TabIndex        =   3
      Top             =   7800
      Width           =   4000
   End
   Begin VB.CommandButton Command1 
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
      Left            =   240
      TabIndex        =   4
      Top             =   7800
      Width           =   4000
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
End Sub

Private Sub getpokebtn_Click()
    
    'Aca se comienza a recibir toda la información de los Pokemon.
    get_pokeinfo
    
End Sub

Public Sub navegar_lista(cant_mostrada As Integer, texto As String)
    
    'Dependiendo del valor de la variable cantidad, se va a mostrar el numero de items deseado por pagina
    'Del item obj se debe hacer otro get para obtener asi la imagen de cada item
    
    Dim html As String
    Dim data_poke() As String
    Dim obj As Object
    Dim objimg As Object
    Dim imgurl As String
    Dim tempurl As String
    Dim count As Integer
    
    Set obj = JSON.parse(texto)
    count = CInt(obj.Item("count"))
    ReDim data_poke(count, count)
    Set httpURL_b = New WinHttp.WinHttpRequest
    
    'Se debe construir un Arreglo que alamacene todos los elementos necesarios, en este caso solo el nombre y la imagen.
    For rec = 1 To 151
    
        'Se debe realizar otro get para obtener la información adicional del Pokemon en cuestión
        'tempurl = obj.Item("results").Item(rec).Item("url")
        
        'httpURL_b.Open "GET", tempurl
        'httpURL_b.Send
        'imgurl = httpURL_b.ResponseText
    
        'If texto = "[]" Then
        '    MsgBox ("No se obtuvo información")
        '    Exit Sub
        'End If
    
        'En caso de recibir información
        Set objimg = JSON.parse(imgurl)
    
        data_poke(rec, 1) = "https://raw.githubusercontent.com/PokeAPI/sprites/master/sprites/pokemon/" & rec & ".png"
        data_poke(rec, 0) = obj.Item("results").Item(rec).Item("name")
    
    Next rec
    
    'MsgBox "Pokemon: " & data_poke(1, 0) & " Imagen: " & data_poke(1, 1)
    
    'Codigo html de ejemplo
    
    html = "<html>" & _
            "<head>" & _
                "<style type='text/css'>" & _
                    "#container_global div{float: left; margin: 10px; background-color: silver;} #container_global div img{width: 100%; height: 100px;}" & _
                "</style>" & _
            "</head>" & _
            "<body>" & _
                "<div id='container_global' style='width:880px; height:600px;'>"
                
    For reg = 1 To cant_mostrada
    
    html = html & "<div style='width:190px; height:200px;'>" & _
                    "<img src=""" & data_poke(reg, 1) & """>" & _
                    "<p align = 'center'><strong>" & data_poke(reg, 0) & "</strong></p>" & _
                    "</div>"
    Next reg
                
    html = html & "</div>" & _
                "</body>" & _
                "<script>alert('hola');<script>" & _
            "</html>"
    
    
    'Se debe construir una lista de items cumpliendo con la cantidad de items que haya establecido el Usuario
    'en el campo "cantfield"
    On Error GoTo ver_error
        
        WebBrowser1.Document.Write html
        WebBrowser1.Refresh
        Exit Sub
        
ver_error:
    MsgBox Err.Description

End Sub

Private Function get_pokeinfo()

    Dim obj As Object
    Dim texto As String
    Dim sinputjson As String
    Dim nombre As String
    Dim url As String
    Dim count As String
    
    Set httpURL = New WinHttp.WinHttpRequest

    cadena = "https://pokeapi.co/api/v2/pokemon?limit=151"
    httpURL.Open "GET", cadena
    httpURL.Send
    texto = httpURL.ResponseText
   
    If texto = "[]" Then
        MsgBox ("No se obtuvo información")
        Exit Function
    End If
    
    If IsNumeric(frm2.cantfield.Text) Then
        Call navegar_lista(CInt(frm2.cantfield.Text), texto)
    Else
        Call navegar_lista(10, texto) 'Valor por defecto
    End If
    
End Function
