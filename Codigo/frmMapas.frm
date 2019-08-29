VERSION 5.00
Begin VB.Form frmMapas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pegardo de zonas"
   ClientHeight    =   4500
   ClientLeft      =   10320
   ClientTop       =   4410
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMapas.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   4650
   Begin VB.Image Image4 
      Height          =   4500
      Left            =   3960
      Picture         =   "frmMapas.frx":FE6F
      Top             =   0
      Width           =   675
   End
   Begin VB.Image Image3 
      Height          =   4500
      Left            =   0
      Picture         =   "frmMapas.frx":15B58
      Top             =   0
      Width           =   675
   End
   Begin VB.Image Image2 
      Height          =   675
      Left            =   750
      Picture         =   "frmMapas.frx":1CE81
      Top             =   0
      Width           =   3150
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   720
      Picture         =   "frmMapas.frx":2421E
      Top             =   3825
      Width           =   3150
   End
End
Attribute VB_Name = "frmMapas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
    SobreX = 1
    SobreY = 95
Call modEdicion.Deshacer_Add("Pegar Selección")
Call pegadoReyarB
'Call modEdicion.Bloquear_Bordes
MapInfo.Changed = 1
DoEvents
Unload Me
End Sub
Private Sub Image2_Click()
    SobreX = 1
    SobreY = 1
Call modEdicion.Deshacer_Add("Pegar Selección")
Call pegadoReyarB
'Call modEdicion.Bloquear_Bordes
MapInfo.Changed = 1
DoEvents
Unload Me
End Sub
Private Sub Image3_Click()
    SobreX = 1
    SobreY = 1
Call modEdicion.Deshacer_Add("Pegar Selección")
Call pegadoReyarB
'Call modEdicion.Bloquear_Bordes
MapInfo.Changed = 1
DoEvents
Unload Me
End Sub
Private Sub Image4_Click()
    SobreX = 92
    SobreY = 1
Call modEdicion.Deshacer_Add("Pegar Selección")
Call pegadoReyarB
'Call modEdicion.Bloquear_Bordes
MapInfo.Changed = 1
DoEvents
Unload Me
End Sub

Private Sub pegadoReyarB()
On Error GoTo Fallo
Call frmUnionAdyacente.Show

    Static UltimoX As Integer
    Static UltimoY As Integer
    'If UltimoX = SobreX And UltimoY = SobreY Then Exit Sub
    UltimoX = SobreX
    UltimoY = SobreY
    Dim X As Integer
    Dim y As Integer
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SobreX
    DeSeleccionOY = SobreY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SobreX, y + SobreY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
             MapData(X + SobreX, y + SobreY) = SeleccionMap(X, y)
        Next
    Next
    Seleccionando = False
    Call DrawMiniMap(True)
    Call modEdicion.Bloquear_Bordes

    Exit Sub

Fallo:
    MsgBox "PegarSeleccion::Error " & Err.Number & " - " & Err.Description
    Call LogError("PegarSeleccion::Error " & Err.Number & " - " & Err.Description)
End Sub


