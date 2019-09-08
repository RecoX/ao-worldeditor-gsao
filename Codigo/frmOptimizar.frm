VERSION 5.00
Begin VB.Form frmOptimizar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Optimizar Mapa"
   ClientHeight    =   4275
   ClientLeft      =   6270
   ClientTop       =   4545
   ClientWidth     =   3600
   Icon            =   "frmOptimizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   3600
   Begin VB.CheckBox chkBloquearTrasladosAngulo 
      Caption         =   "Quitar Traslados y Bloqueos en angulos de los mapas"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CheckBox chkBloquearArbolesEtc 
      Caption         =   "Bloquear Arboles, Carteles, Foros y Yacimientos"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkMapearArbolesEtc 
      Caption         =   "Mapear Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTodoBordes 
      Caption         =   "Quitar NPCs, Objetos y Traslados en los Bordes Exteriores"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrigTras 
      Caption         =   "Quitar Trigger's en Traslados"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrigBloq 
      Caption         =   "Quitar Trigger's Bloqueados"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrans 
      Caption         =   "Quitar Translados Bloqueados"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin GSZAOWorldEditor.lvButtons_H cOptimizar 
      Default         =   -1  'True
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      Caption         =   "&Optimizar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12648384
   End
   Begin GSZAOWorldEditor.lvButtons_H cCancelar 
      Height          =   735
      Left            =   1920
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      Caption         =   "&Cancelar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmOptimizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Optimizar()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************

Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

' Quita Traslados Bloqueados
' Quita Trigger's Bloqueados
' Quita Trigger's en Traslados
' Quita NPCs, Objetos y Traslados en los Bordes Exteriores
' Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa

modEdicion.Deshacer_Add "Aplicar Optimizacion del Mapa" ' Hago deshacer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        ' ** Quitar NPCs, Objetos y Traslados en los Bordes Exteriores
        If (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) And chkQuitarTodoBordes.Value = 1 Then
             'Quitar NPCs
            If MapData(X, Y).NPCIndex > 0 Then
                EraseChar MapData(X, Y).CharIndex
                MapData(X, Y).NPCIndex = 0
            End If
            ' Quitar Objetos
            MapData(X, Y).OBJInfo.objindex = 0
            MapData(X, Y).OBJInfo.Amount = 0
            MapData(X, Y).ObjGrh.grh_index = 0
            ' Quitar Traslados
            MapData(X, Y).TileExit.Map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0
            ' Quitar Triggers
            MapData(X, Y).Trigger = 0
        End If
        ' ** Quitar Traslados y Triggers en Bloqueo
        If MapData(X, Y).Blocked = 1 Then
            If MapData(X, Y).TileExit.Map > 0 And chkQuitarTrans.Value = 1 Then ' Quita Translado Bloqueado
                MapData(X, Y).TileExit.Map = 0
                MapData(X, Y).TileExit.Y = 0
                MapData(X, Y).TileExit.X = 0
            ElseIf MapData(X, Y).Trigger > 0 And chkQuitarTrigBloq.Value = 1 Then ' Quita Trigger Bloqueado
                MapData(X, Y).Trigger = 0
            End If
        End If
        ' ** Quitar Triggers en Translado
        If MapData(X, Y).TileExit.Map > 0 And chkQuitarTrigTras.Value = 1 Then
            If MapData(X, Y).Trigger > 0 Then ' Quita Trigger en Translado
                MapData(X, Y).Trigger = 0
            End If
        End If
        ' ** Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
        If MapData(X, Y).OBJInfo.objindex > 0 And (chkMapearArbolesEtc.Value = 1 Or chkBloquearArbolesEtc.Value = 1) Then
            Select Case ObjData(MapData(X, Y).OBJInfo.objindex).ObjType
                Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                    If MapData(X, Y).Graphic(3).grh_index <> MapData(X, Y).ObjGrh.grh_index And chkMapearArbolesEtc.Value = 1 Then MapData(X, Y).Graphic(3) = MapData(X, Y).ObjGrh
                    If chkBloquearArbolesEtc.Value = 1 And MapData(X, Y).Blocked = 0 Then MapData(X, Y).Blocked = 1
            End Select
        End If
        ' ** Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
    Next X
Next Y
    If chkBloquearTrasladosAngulo.Value = 1 Then
             ' Quitar Bloqueso en angulos
            MapData(9, 7).Blocked = 0
            MapData(92, 7).Blocked = 0
            MapData(9, 94).Blocked = 0
            MapData(92, 94).Blocked = 0

            ' Quitar Traslados en angulos
            MapData(9, 7).TileExit.Map = 0
            MapData(9, 7).TileExit.X = 0
            MapData(9, 7).TileExit.Y = 0
            
            MapData(92, 7).TileExit.Map = 0
            MapData(92, 7).TileExit.X = 0
            MapData(92, 7).TileExit.Y = 0
            
            MapData(9, 94).TileExit.Map = 0
            MapData(9, 94).TileExit.X = 0
            MapData(9, 94).TileExit.Y = 0
            
            MapData(92, 94).TileExit.Map = 0
            MapData(92, 94).TileExit.X = 0
            MapData(92, 94).TileExit.Y = 0
            
            MapInfo.Changed = 1
            End If
'Set changed flag
MapInfo.Changed = 1

End Sub

Private Sub cCancelar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/09/06
'*************************************************
Unload Me
End Sub



Private Sub cOptimizar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/09/06
'*************************************************
Call Optimizar
MapInfo.Changed = 1
DoEvents

Unload Me
End Sub


