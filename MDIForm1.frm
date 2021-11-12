VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   4560
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   6720
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu menuaa 
      Caption         =   "Consulta de datos"
      Index           =   0
   End
   Begin VB.Menu menupokeapi 
      Caption         =   "PokeAPI"
      Index           =   1
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub menuaa_Click(Index As Integer)
    frm1.Show
End Sub

Private Sub menupokeapi_Click(Index As Integer)
    frm2.Show
End Sub
