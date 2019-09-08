VERSION 5.00
Begin VB.Form frmAutoGuardarMapa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guardar Mapa Automaticamente"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "frmAutoGuardarMapa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmAutoGuardarMapa.frx":628A
   ScaleHeight     =   2895
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin GSZAOWorldEditor.lvButtons_H cmdAceptar 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "&Aceptar y Aplicar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12648384
   End
   Begin GSZAOWorldEditor.lvButtons_H cmdCerrar 
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Cerrar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.ComboBox cmbMinutos 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmAutoGuardarMapa.frx":160F9
      Left            =   1440
      List            =   "frmAutoGuardarMapa.frx":16112
      TabIndex        =   1
      Text            =   "10"
      Top             =   1440
      Width           =   1215
   End
   Begin GSZAOWorldEditor.lvButtons_H cmdDetener 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Detener"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632319
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   135
      X2              =   4440
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "minutos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Indique cada cuantos Minutos desea que se Guarde Automaticamente el Mapa con el que trabaja:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   4440
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "frmAutoGuardarMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Option Explicit

Private Sub cmdAceptar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If IsNumeric(cmbMinutos.Text) = False Then
    MsgBox "Los minutos deben ingresarse de forma númerica.", vbCritical + vbOKOnly
    Exit Sub
ElseIf Val(cmbMinutos.Text) < 5 Or Val(cmbMinutos.Text) > 120 Then
    MsgBox "Los minutos ingresados son invalidos." & vbCrLf & "Solo estan permitidos los valores de entre 5 y 120 minutos inclusive.", vbCritical + vbOKOnly
    Exit Sub
End If
bAutoGuardarMapa = Val(cmbMinutos.Text)
bAutoGuardarMapaCount = 0
frmMain.TimAutoGuardarMapa.Enabled = True
frmMain.mnuAutoGuardarMapas.Checked = True
Unload Me
End Sub

Private Sub cmdCerrar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Unload Me
End Sub

Private Sub cmdDetener_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmMain.TimAutoGuardarMapa.Enabled = False
frmMain.mnuAutoGuardarMapas.Checked = False
Unload Me
End Sub

Private Sub Form_Load()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
cmbMinutos.Text = bAutoGuardarMapa
End Sub
