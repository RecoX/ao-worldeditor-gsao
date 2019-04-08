Attribute VB_Name = "modImageProcess"
'World Grid Maker
'Copyright (C) 2012 GS-Zone
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'You can contact me at:
'info@gs-zone.org

Option Explicit
'
'Requires a reference to:
'   Microsoft Windows Image Acquisition Library v2.0
'

Private Const TIFF_LZW As String = "LZW"
Private Const TIFF_RLE As String = "RLE"       'Pixel Depth must be 1.
Private Const TIFF_CCITT3 As String = "CCITT3" 'Pixel Depth must be 1.
Private Const TIFF_CCITT4 As String = "CCITT4" 'Pixel Depth must be 1.
Private Const TIFF_Uncompressed As String = "Uncompressed"

Public Sub WIAImageProcess( _
    ByVal InFileName As String, _
    ByVal OutFileName As String, _
    ByVal OutFormat As String, _
    Optional ByVal Quality As Integer = 100, _
    Optional ByVal Compression As String = TIFF_LZW)

    Dim Img As WIA.ImageFile
    Dim ImgProc As WIA.ImageProcess

    Set Img = New WIA.ImageFile
    Img.LoadFile InFileName
    Set ImgProc = New WIA.ImageProcess
    With ImgProc.Filters
        .Add ImgProc.FilterInfos("Convert").FilterID
        .Item(1).Properties("FormatID").Value = OutFormat
        If OutFormat = wiaFormatJPEG Then
            .Item(1).Properties("Quality").Value = Quality
        ElseIf OutFormat = wiaFormatTIFF Then
            .Item(1).Properties("Compression").Value = Compression
        End If
    End With
    Set Img = ImgProc.Apply(Img)

    On Error Resume Next
    Kill OutFileName
    On Error GoTo 0
    Img.SaveFile OutFileName
End Sub

'    ImgConvert "sample.bmp", "sample.jpg", wiaFormatJPEG, 70
'    ImgConvert "sample.bmp", "sample.gif", wiaFormatGIF
'    ImgConvert "sample.bmp", "sample.png", wiaFormatPNG
'    ImgConvert "sample.bmp", "sample.tif", wiaFormatTIFF, , TIFF_Uncompressed

