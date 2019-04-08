Attribute VB_Name = "modGameIni"
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
' modGameIni
'
' @remarks Operaciones de Cabezera y inicio.con
' @author unkwown
' @version 0.0.01
' @date 20060520

Option Explicit

' GSZAO - Archivos de configuraci�n!
Public Const fAOSetup = "AOSetup.init"
Public Const fConfigInit = "Config.init"

' GSZAO - Las variables de path se definen una sola vez (Ver Sub InitFilePaths())
Public DirGraficos As String
Public DirSound As String
Public DirMidi As String
Public DirMapas As String
Public DirExtras As String
Public DirCursores As String
Public DirGUI As String
Public DirButtons As String

Public Const nDirINIT = "\INIT\" ' Directorio INIT
Public sPathINIT As String ' Path de INITs

' Variable solo WorldEditor
Public DirDat As String
Public DirMapIndex As String
Public Const nDirMAPINDEX = "\MAPINDEX\" ' Directorio de MapIndex
Public Const nDirGRAFICOS = "\GRAFICOS\" ' Directorio de Graficos
Public Const nDirDAT = "\DATS\" ' Directorio de Dats
Public Const nDirMIDI = "\MIDI\" ' Directorio de Musica
Public WorldEditorIni As String ' Archivo de configuraci�n del WorldEditor!
Public WorldEditorQuickSup As String ' Archivo de configuraci�n con las superficies de acceso rapido personales
Public IniPath As String ' Directorio del WorldEditor!
Public bGraficosAO As Boolean
Public bAutoPantalla As Boolean ' Determinar el tama�o de trabajo automaticamente!
Public bBuscarErroresEnGrhIndex As Boolean ' �Buscar errores?!

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tConfigInit
    ' Opciones
    MostrarTips As Byte         ' Activa o desactiva la muestra de tips
    NumParticulas As Integer    ' Numero de particulas
    IndiceGraficos As String    ' Archivo de Indices de Graficos
    
    ' Usuario
    Nombre As String            ' Nombre de usuario
    Password As String          ' Contrase�a del usuario
    Recordar As Byte            ' Activado el recordar!
    
    ' Directorio
    DirMultimedia As String     ' Directorio de multimedia
    DirMapas As String          ' Directorio de mapas
    DirGraficos As String       ' Directorio de graficos
    DirFotos As String          ' Directorio de fotos
    DirExtras As String         ' Directorio de extras (dentro de inits)
    DirSonidos As String        ' Directorio de sonidos (dentro de multimedia)
    DirMusicas As String        ' Directorio de musicas (dentro de multimedia)
    DirParticulas As String     ' Directorio de particulas (dentro de graficos)
    DirCursores As String       ' Directorio de cursores (dentro de graficos)
    DirGUI As String            ' Directorio del GUI (dentro de graficos)
    DirBotones As String        ' Directorio de botones (dentro de GUI)
    DirFrags As String          ' Directorio de frags (dentro de fotos)
    DirMuertes As String        ' Directorio de muertes (dentro de fotos)
End Type

Public Type tAOSetup
    ' VIDEO
    bVertex     As Byte     ' GSZAO - Cambia el Vortex de dibujado
    bVSync      As Boolean  ' GSZAO - Utiliza Sincronizaci�n Vertical (VSync)
    bDinamic    As Boolean  ' Utilizar carga Dinamica de Graficos o Estatica
    byMemory    As Byte     ' Uso maximo de memoria para la carga Dinamica (exclusivamente)

    ' SONIDO
    bNoMusic    As Boolean  ' Jugar sin Musica
    bNoSound    As Boolean  ' Jugar sin Sonidos
    bNoSoundEffects As Boolean  ' Jugar sin Efectos de sonido (basicamente, sonido que viene de la izquierda y de la derecha)
    lMusicVolume As Long ' Volumen de la Musica
    lSoundVolume As Long ' Volumen de los Sonidos
    
    ' SCREENSHOTS
    bActive     As Boolean  ' Activa el modo de screenshots
    bDie        As Boolean  ' Obtiene una screenshot al morir (si bActive = True)
    bKill       As Boolean  ' Obtiene una screenshot al matar (si bActive = True)
    byMurderedLevel As Byte ' La screenshot al matar depende del nivel de la victima (si bActive = True)
    
    ' CLAN
    bGuildNews  As Boolean      ' Mostrar Noticias del Clan al inicio
    bGldMsgConsole As Boolean   ' Activa los Dialogos de Clan
    bCantMsgs   As Byte         ' Establece el maximo de mensajes de Clan en pantalla
End Type

Public MiCabecera As tCabecera
Public ClientConfigInit As tConfigInit
Public ClientAOSetup As tAOSetup

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
'**************************************************************
'Author: Unknown
'Last Modify Date: 04/08/2012 - ^[GS]^
'**************************************************************
    Cabecera.Desc = "GS-Zone Argentum Online MOD - Copyright GS-Zone 2012 - info@gs-zone.org - Original by Pablo Marquez " ' GSZAO
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
    
End Sub

Public Function LeerConfigInit() As tConfigInit
'**************************************************************
'Author: ^[GS]^
'Last Modify Date: 04/08/2012 - ^[GS]^
'**************************************************************
On Local Error Resume Next

    Dim N As Integer
    Dim ConfigInit As tConfigInit
    N = FreeFile
    Open sPathINIT & fConfigInit For Binary As #N
        Get #N, , MiCabecera
        Get #N, , ConfigInit
    Close #N
    
    LeerConfigInit = ConfigInit
    
End Function

Public Sub InitGraphicsFile()
'*************************************************
'Author: ^[GS]^
'Last modified: 04/08/2012 - ^[GS]^
'*************************************************
    If InStr(1, ClientConfigInit.IndiceGraficos, "Graficos") Then
        GraphicsFile = ClientConfigInit.IndiceGraficos
    Else
        GraphicsFile = "Graficos.ind"
    End If
End Sub

Public Sub LoadClientAOSetup()
'**************************************************************
'Author: ^[GS]^
'Last Modification: 04/08/2012 - ^[GS]^
'**************************************************************
    Dim fHandle As Integer
    
    ' Por default
    ClientAOSetup.bDinamic = True
    ClientAOSetup.bVertex = 0 ' software
    ClientAOSetup.bVSync = False
    If FileExist(sPathINIT & fAOSetup, vbArchive) Then
        fHandle = FreeFile
        Open sPathINIT & fAOSetup For Binary Access Read Lock Write As fHandle
            Get fHandle, , ClientAOSetup
        Close fHandle
    End If
End Sub


Public Function NombrePC() As String
'**************************************************************
'Author: ^[GS]^
'Last Modification: 01/04/2013 - ^[GS]^
'**************************************************************

    Dim dwLen As Long
    Dim strString As String
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    GetComputerName strString, dwLen
    strString = Left(strString, dwLen)
    NombrePC = strString
    
End Function
