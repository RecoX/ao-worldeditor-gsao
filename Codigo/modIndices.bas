Attribute VB_Name = "modIndices"
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

''
' modIndices
'
' @remarks Funciones Especificas al Trabajo con Indices
' @author gshaxor@gmail.com
' @version 0.1.05
' @date 20060530

Option Explicit

' Carga los indices de Superficie
'

Public Sub CargarIndicesSuperficie()
'*************************************************
'Author: ^[GS]^
'Last modified: 04/08/2012 - ^[GS]^
'*************************************************

On Error GoTo Fallo
    If Not FileExist(DirMapIndex & "GrhIndex.ini", vbArchive) Then
        MsgBox "Falta el archivo '" & DirMapIndex & "GrhIndex.ini'", vbCritical
        End
    End If
    
    Dim Leer As New clsIniReader
    Dim i As Integer
    Dim t As Integer
    
    Leer.Initialize DirMapIndex & "GrhIndex.ini"
    
    MaxSup = Leer.GetValue("INIT", "Referencias")
    
    ReDim SupData(MaxSup) As SupData
    frmMain.lListado(0).Clear
    
    For i = 0 To MaxSup
        SupData(i).Name = Leer.GetValue("REFERENCIA" & i, "Nombre")
        SupData(i).Grh = Val(Leer.GetValue("REFERENCIA" & i, "GrhIndice"))
        SupData(i).Width = Val(Leer.GetValue("REFERENCIA" & i, "Ancho"))
        SupData(i).Height = Val(Leer.GetValue("REFERENCIA" & i, "Alto"))
        SupData(i).Block = IIf(Val(Leer.GetValue("REFERENCIA" & i, "Bloquear")) = 1, True, False)
        SupData(i).Capa = Val(Leer.GetValue("REFERENCIA" & i, "Capa"))
        If LenB(SupData(i).Name) <> 0 Then
            If bBuscarErroresEnGrhIndex = True Then ' GSZAO
                For t = 0 To i - 1
                    If LenB(SupData(t).Name) <> 0 Then
                        If SupData(t).Grh = SupData(i).Grh Then ' ¿usa mismo grh?!
                            If SupData(t).Width = SupData(i).Width And SupData(t).Height = SupData(i).Height Then ' mismo tamaño!
                                MsgBox "El indice " & i & " (" & SupData(i).Name & ") tiene repetido el Grh con el indice " & t & " (" & SupData(t).Name & ")", vbInformation + vbOKOnly
                            End If
                        End If
                    End If
                Next
            End If
            frmMain.lListado(0).AddItem SupData(i).Name & " - #" & i
        End If
    Next i
    
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el indice " & i & " de " & DirMapIndex & "GrhIndex.ini" & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
End Sub

''
' Carga los indices de Objetos
'

Public Sub CargarIndicesOBJ()
'*************************************************
'Author: ^[GS]^
'Last modified: 13/02/2014
'*************************************************

On Error GoTo Fallo
    If FileExist(DirDat & "\OBJ.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'OBJ.dat' en " & DirDat, vbCritical
        End
    End If
    Dim strTipo As String
    Dim Obj As Integer
    Dim Leer As New clsIniReader
    Call Leer.Initialize(DirDat & "\OBJ.dat")
    frmMain.lListado(3).Clear
    NumOBJs = Val(Leer.GetValue("INIT", "NumOBJs"))
    ReDim ObjData(1 To NumOBJs) As ObjData
    For Obj = 1 To NumOBJs
        frmCargando.X.Caption = "Cargando Datos de Objetos..." & Obj & "/" & NumOBJs
        DoEvents
        ObjData(Obj).Name = Leer.GetValue("OBJ" & Obj, "Name")
        ObjData(Obj).grh_index = Val(Leer.GetValue("OBJ" & Obj, "GrhIndex"))
        ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
        ObjData(Obj).Ropaje = Val(Leer.GetValue("OBJ" & Obj, "NumRopaje"))
        ObjData(Obj).Info = Leer.GetValue("OBJ" & Obj, "Info")
        ObjData(Obj).WeaponAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
        ObjData(Obj).Texto = Leer.GetValue("OBJ" & Obj, "Texto")
        ObjData(Obj).GrhSecundario = Val(Leer.GetValue("OBJ" & Obj, "GrhSec"))
        
        strTipo = GetObjType(ObjData(Obj).ObjType)
    
        frmMain.lListado(3).AddItem "[" & strTipo & "] " & ObjData(Obj).Name & " - #" & Obj
    Next Obj
    Exit Sub
    
Fallo:
    MsgBox "Error al intentar cargar el Objteto " & Obj & " de OBJ.dat en " & DirDat & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Triggers
'

Public Sub CargarIndicesTriggers()
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012 - ^[GS]^
'*************************************************
On Error GoTo Fallo
    If FileExist(DirMapIndex & "Triggers.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'Triggers.ini' en " & DirMapIndex, vbCritical
        End
    End If
    Dim NumT As Integer
    Dim t As Integer
    Dim Leer As New clsIniReader
    
    frmMain.lListado(4).Clear
    frmMain.lListado(4).AddItem "Sin Trigger - #0"
    
    Call Leer.Initialize(DirMapIndex & "Triggers.ini")
    NumT = Val(Leer.GetValue("INIT", "NumTriggers"))
    
    If NumT < 1 Then Exit Sub ' GSZAO
    ReDim TriggerData(NumT) As tTriggerData ' GSZAO
    Dim sColor As String
    
    For t = 1 To NumT
        Select Case t
            Case 1
                sColor = " [Rojo]" ' 255,0,0
            Case 2
                sColor = " [Verde]" ' 0,255,0
            Case 3
                sColor = " [Azul]" ' 0,0,255
            Case 4
                sColor = " [Celeste]" ' 0,255,255
            Case 5
                sColor = " [Naranja]" ' 255,64,0
            Case 6
                sColor = " [Rozado]" ' 255,128,255
            Case Else
                sColor = " [Amarillo]" ' 255,255,0
        End Select
        TriggerData(t).id = t
        TriggerData(t).Name = Leer.GetValue("Trig" & t, "Name")
        TriggerData(t).Desc = Leer.GetValue("Trig" & t, "Desc") ' GSZAO
        frmMain.lListado(4).AddItem TriggerData(t).Name & " - #" & TriggerData(t).id & sColor
    Next t

Exit Sub

Fallo:
    MsgBox "Error al intentar cargar el Trigger " & t & " de Triggers.ini en " & DirMapIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Cuerpos
'

Public Sub CargarIndicesDeCuerpos()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo Fallo
    If Not FileExist(sPathINIT & "Cuerpos.ind", vbArchive) Then
        MsgBox "Falta el archivo 'Cuerpos.ind' en " & sPathINIT, vbCritical
        End
    End If
    
    Dim N As Integer
    Dim i As Integer
    
    N = FreeFile
    Open sPathINIT & "Cuerpos.ind" For Binary Access Read As #N
        'cabecera
        Get #N, , MiCabecera
        'num de cabezas
        Get #N, , NumBodies
        
        'Resize array
        ReDim BodyData(1 To NumBodies) As tBodyData
        ReDim MisCuerpos(1 To NumBodies) As tIndiceCuerpo
        
        For i = 1 To NumBodies
            Get #N, , MisCuerpos(i)
            
            Grh_Initialize BodyData(i).Walk(1), MisCuerpos(i).Body(1), , , 0
            Grh_Initialize BodyData(i).Walk(2), MisCuerpos(i).Body(2), , , 0
            Grh_Initialize BodyData(i).Walk(3), MisCuerpos(i).Body(3), , , 0
            Grh_Initialize BodyData(i).Walk(4), MisCuerpos(i).Body(4), , , 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        Next i
    Close #N
    
Exit Sub

Fallo:
    MsgBox "Error al intentar cargar el Cuerpo " & i & " de Cuerpos.ind en " & sPathINIT & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Cabezas
'

Public Sub CargarIndicesDeCabezas()
On Error GoTo Fallo
    If Not FileExist(sPathINIT & "Cabezas.ind", vbArchive) Then
        MsgBox "Falta el archivo 'Cabezas.ind' en " & sPathINIT, vbCritical
        End
    End If
    
    Dim N As Integer
    Dim i As Long
    Dim MisCabezas() As tIndiceCabeza
    
    N = FreeFile()
    
    Open sPathINIT & "Cabezas.ind" For Binary Access Read As #N
        'cabecera
        Get #N, , MiCabecera
        'num de cabezas
        Get #N, , Numheads
        'Resize array
        ReDim HeadData(0 To Numheads) As tHeadData
        ReDim MisCabezas(0 To Numheads) As tIndiceCabeza
        
        For i = 1 To Numheads
            Get #N, , MisCabezas(i)
            
            If MisCabezas(i).Head(1) Then
                Call Grh_Initialize(HeadData(i).Head(1), MisCabezas(i).Head(1), , , 0)
                Call Grh_Initialize(HeadData(i).Head(2), MisCabezas(i).Head(2), , , 0)
                Call Grh_Initialize(HeadData(i).Head(3), MisCabezas(i).Head(3), , , 0)
                Call Grh_Initialize(HeadData(i).Head(4), MisCabezas(i).Head(4), , , 0)
            End If
        Next i
    Close #N
    
Exit Sub

Fallo:
    MsgBox "Error al intentar cargar la Cabeza " & i & " de Cabezas.ind en " & sPathINIT & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub


''
' Carga los indices de NPCs
'

Public Sub CargarIndicesNPC()
'*************************************************
'Author: ^[GS]^
'Last modified: 25/08/2012 - ^[GS]^
'*************************************************
On Error Resume Next
'On Error GoTo Fallo
    If FileExist(DirDat & "\NPCs.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'NPCs.dat' en " & DirDat, vbCritical
        End
    End If
    Dim NPC As Integer
    Dim Leer As New clsIniReader
    frmMain.lListado(1).Clear
    frmMain.lListado(2).Clear
    Call Leer.Initialize(DirDat & "\NPCs.dat")
    NumNPCs = Val(Leer.GetValue("INIT", "NumNPCs"))

    ReDim NpcData(NumNPCs) As NpcData

    For NPC = 1 To NumNPCs
        With NpcData(NPC)
            .Name = Leer.GetValue("NPC" & NPC, "Name")
            
            .Body = Val(Leer.GetValue("NPC" & NPC, "Body"))
            .Head = Val(Leer.GetValue("NPC" & NPC, "Head"))
            .Heading = Val(Leer.GetValue("NPC" & NPC, "Heading"))
            .Hostile = CBool(Val(Leer.GetValue("NPC" & NPC, "Hostile")))
            
            If LenB(.Name) <> 0 Or .Body <> 0 Then
                If .Hostile = True Then
                    frmMain.lListado(1).AddItem .Name & " [HOSTIL] - #" & NPC
                Else
                    frmMain.lListado(1).AddItem .Name & " - #" & NPC
                End If
            End If
            'If LenB(NpcData(NPC).name) <> 0 Then frmMain.lListado(1).AddItem NpcData(NPC).name & " - #" & NPC
        End With
    Next

    Exit Sub
    
Fallo:
    MsgBox "Error al intentar cargar el NPC " & NPC & " de 'NPCs.dat' en " & DirDat & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub


' GSZAO - Tipos de Objetos
Function GetObjType(ByVal ObjType As Integer) As String
'*************************************************
'Author: ^[GS]^
'Last modified: 13/02/2014 - ^[GS]^
'*************************************************

        Select Case ObjType
            Case 1
                GetObjType = "Comida"
            Case 2
                GetObjType = "Arma"
            Case 3
                GetObjType = "Armadura"
            Case 4
                GetObjType = "Arbol"
            Case 5
                GetObjType = "Dinero"
            Case 6
                GetObjType = "Puerta"
            Case 7
                GetObjType = "Contenedor"
            Case 8
                GetObjType = "Cartel"
            Case 9
                GetObjType = "Llave"
            Case 10
                GetObjType = "Foro"
            Case 11
                GetObjType = "Pocion"
            Case 12
                GetObjType = "Libro"
            Case 13
                GetObjType = "Bebida"
            Case 14
                GetObjType = "Leña"
            Case 15
                GetObjType = "Fogata"
            Case 16
                GetObjType = "Escudo"
            Case 17
                GetObjType = "Casco"
            Case 18
                GetObjType = "Anillo"
            Case 19
                GetObjType = "Teletransporte"
        Case 20
            GetObjType = "Mueble"
        Case 21
            GetObjType = "Joya"
        Case 22
            GetObjType = "Yacimiento"
        Case 23
            GetObjType = "Metal"
        Case 24
            GetObjType = "Pergamino"
        Case 25
            GetObjType = "Aura"
        Case 26
            GetObjType = "Instrumento"
        Case 27
            GetObjType = "Yunque"
        Case 28
            GetObjType = "Fragua"
        Case 29
            GetObjType = "Gema"
        Case 30
            GetObjType = "Flor"
        Case 31
            GetObjType = "Barco"
        Case 32
            GetObjType = "Flecha"
        Case 33
            GetObjType = "Botella vacia"
        Case 34
            GetObjType = "Botella llena"
        Case 35
            GetObjType = "Mancha"
        Case 36
            GetObjType = "Arbol Elfico"
        Case 37
            GetObjType = "Mochila"
        Case 38
            GetObjType = "Cardumen"
        Case 39
            GetObjType = "Pasaje"
        Case 40
            GetObjType = "Obj. destruible"
        Case 41
            GetObjType = "Matrimonio"
        Case Else
            GetObjType = "??"
    End Select
End Function


