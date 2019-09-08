VERSION 5.00
Begin VB.Form frmcopiarmapa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiado de Zonas"
   ClientHeight    =   4485
   ClientLeft      =   10380
   ClientTop       =   4305
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmcopiarmapa.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   4680
   Begin VB.CommandButton iralmapa 
      Caption         =   "Ir al Mapa"
      BeginProperty Font 
         Name            =   "Morpheus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Mapa 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   4500
      Left            =   0
      Picture         =   "frmcopiarmapa.frx":FE6F
      Top             =   0
      Width           =   750
   End
   Begin VB.Image Image3 
      Height          =   4500
      Left            =   3960
      Picture         =   "frmcopiarmapa.frx":171FF
      Top             =   0
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   670
      Left            =   830
      Picture         =   "frmcopiarmapa.frx":1E4C7
      Top             =   3830
      Width           =   3150
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   720
      Picture         =   "frmcopiarmapa.frx":25A64
      Top             =   0
      Width           =   3150
   End
End
Attribute VB_Name = "frmcopiarmapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
'COPIADO ZONA AZUL "ARRIBA"

On Error Resume Next

Dim X As Integer
Dim y As Integer

' ARRIBA
Mapa.Text = 0
y = 7
For X = (9 + 1) To (92 - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa.Text = MapData(X, y).TileExit.Map
            Exit For
        End If
Next

Call modEdicion.Quitar_Translados_ReyarB
Call modEdicion.Bloqueo_Todos(0)


    SeleccionIX = 1
    SeleccionFX = 100
    SeleccionIY = 8
    SeleccionFY = 14
Call CopiarSeleccionMapa
frmMapas.Image1.Visible = True
frmMapas.Image2.Visible = False
frmMapas.Image3.Visible = False
frmMapas.Image4.Visible = False
MapInfo.Changed = 1
UserPos.X = 9
UserPos.y = 100
End Sub

Private Sub Image2_Click()
'COPIADO ZONA VERDE "ABAJO"

On Error Resume Next

Dim X As Integer
Dim y As Integer

' ABAJO
Mapa.Text = 0
y = 94
For X = (9 + 1) To (92 - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa.Text = MapData(X, y).TileExit.Map
            Exit For
        End If
Next

Call modEdicion.Quitar_Translados_ReyarB
Call modEdicion.Bloqueo_Todos(0)

    SeleccionIX = 1
    SeleccionFX = 100
    SeleccionIY = 87
    SeleccionFY = 93
Call CopiarSeleccionMapa
frmMapas.Image1.Visible = False
frmMapas.Image2.Visible = True
frmMapas.Image3.Visible = False
frmMapas.Image4.Visible = False
UserPos.X = 9
UserPos.y = 15
End Sub

Private Sub Image3_Click()
'COPIADO ZONA AMARILLA "DERECHA"

Dim X As Integer
Dim y As Integer
'DERECHA
Mapa.Text = 0
X = 92
For y = (7 + 1) To (94 - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa.Text = MapData(X, y).TileExit.Map
            Exit For
        End If
Next
Call modEdicion.Quitar_Translados_ReyarB
Call modEdicion.Bloqueo_Todos(0)


    SeleccionIX = 83
    SeleccionFX = 91
    SeleccionIY = 1
    SeleccionFY = 100
    Call CopiarSeleccionMapa
frmMapas.Image1.Visible = False
frmMapas.Image2.Visible = False
frmMapas.Image3.Visible = True
frmMapas.Image4.Visible = False
UserPos.X = 9
UserPos.y = 15
End Sub

Private Sub Image4_Click()
'COPIADO ZONA ROJA "IZQUIERDA"

On Error Resume Next

Dim X As Integer
Dim y As Integer
'DERECHA

Mapa.Text = 0
X = 9
For y = (7 + 1) To (94 - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa.Text = MapData(X, y).TileExit.Map
            Exit For
        End If
Next

Call modEdicion.Quitar_Translados_ReyarB
Call modEdicion.Bloqueo_Todos(0)

    SeleccionIX = 10
    SeleccionFX = 18
    SeleccionIY = 1
    SeleccionFY = 100
Call CopiarSeleccionMapa
frmMapas.Image1.Visible = False
frmMapas.Image2.Visible = False
frmMapas.Image3.Visible = False
frmMapas.Image4.Visible = True
UserPos.X = 85
UserPos.y = 15
End Sub
Public Sub CopiarSeleccionMapa()
'*************************************************
'Author: ReyarB
'Last modified: 01/20/2018
'*************************************************

On Error GoTo Fallo

    'podria usar copy mem , pero por las dudas no XD
    Dim X As Integer
    Dim y As Integer
    'Seleccionando = False
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    ReDim SeleccionMap(SeleccionAncho, SeleccionAlto) As MapBlock
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
                If MapData(X + SeleccionIX, y + SeleccionIY).NPCIndex > 0 Then
                    EraseChar MapData(X + SeleccionIX, y + SeleccionIY).CharIndex
                    MapData(X + SeleccionIX, y + SeleccionIY).NPCIndex = 0
                End If
                    MapData(X + SeleccionIX, y + SeleccionIY).OBJInfo.objindex = 0
                    MapData(X + SeleccionIX, y + SeleccionIY).OBJInfo.Amount = 0
                    MapData(X + SeleccionIX, y + SeleccionIY).ObjGrh.grh_index = 0
                    ' Quitar Translados
                    MapData(X + SeleccionIX, y + SeleccionIY).TileExit.Map = 0
                    MapData(X + SeleccionIX, y + SeleccionIY).TileExit.X = 0
                    MapData(X + SeleccionIX, y + SeleccionIY).TileExit.y = 0
                 ' Quitar Triggers
                    MapData(X + SeleccionIX, y + SeleccionIY).Trigger = 0
            SeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
frmcopiarmapa.Mapa.Visible = True
frmcopiarmapa.iralmapa.Visible = True

frmcopiarmapa.Image1.Visible = False
frmcopiarmapa.Image2.Visible = False
frmcopiarmapa.Image3.Visible = False
frmcopiarmapa.Image4.Visible = False

Exit Sub

Fallo:
    MsgBox "CopiarSeleccion::Error " & Err.Number & " - " & Err.Description
    Call LogError("CopiarSeleccion::Error " & Err.Number & " - " & Err.Description)

End Sub

Private Sub iralmapa_Click()
'Public Sub CambioMapa()

'*************************************************
'Author: ^[ReyarB]^
'Last modified: 01/01/2018 - ^[GS]^
'*************************************************

On Error GoTo Fallo
    ' Selecciones
    Seleccionando = False
    SeleccionIX = 0
    SeleccionIY = 0
    SeleccionFX = 0
    SeleccionFY = 0
    ' Traslados
    Dim tTrans As WorldPos
 
    If tTrans.Map = 0 Then
        If LenB(frmMain.Dialog.FileName) > 0 Then
            If FileExist(PATH_Save & NameMap_Save & Mapa & ".map", vbArchive) = True Then
                Call modMapIO.NuevoMapa
                frmMain.Dialog.FileName = PATH_Save & NameMap_Save & Mapa & ".map"
                modMapIO.AbrirMapa frmMain.Dialog.FileName, MapData
                
                If WalkMode = True Then
                    MoveCharbyPos UserCharIndex, UserPos.X, UserPos.y
                    CharList(UserCharIndex).Heading = SOUTH
                End If
                frmMain.mnuReAbrirMapa.Enabled = True
            End If
        End If
    End If
frmcopiarmapa.Mapa.Visible = False
frmcopiarmapa.iralmapa.Visible = False
frmcopiarmapa.Image1.Visible = True
frmcopiarmapa.Image2.Visible = True
frmcopiarmapa.Image3.Visible = True
frmcopiarmapa.Image4.Visible = True
frmMapas.Show
DoEvents
Unload Me
    Exit Sub

Fallo:
    MsgBox "DobleClick::Error " & Err.Number & " - " & Err.Description

End Sub


Private Sub Mapa_Change()
If Mapa.Text = 0 Then
    iralmapa.Caption = "Salir"
    MapInfo.Changed = 0
    Else
    iralmapa.Caption = "Ir al Mapa"
End If
End Sub
