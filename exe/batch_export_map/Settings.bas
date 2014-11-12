Attribute VB_Name = "Settings"
Option Explicit

Public outputFormat As String
Public embedFonts As Boolean
Public convMarkPoly As Boolean
Public mapCompress As Boolean
Public compType As String
Public pictSymb As String
Public colorMode As String
Public traspCol As Boolean
Public dpi As Integer
Public imageQuality As String
Public exportError As Boolean
Public closeApp As Boolean
Public agpSelFileArray()
Public includeSub As Boolean
Public outputMaps()
Public numOutput As Integer
Public layersAttrib As String
Public expGeoInfo As String


Public Sub SaveAsDefaults()

Dim filenum As Integer

If Export_Map.txb_dpi.Text = "" Then
    MsgBox "ATTENTION: you've to set DPI resolution!", vbExclamation, "Warning"
    Exit Sub
End If


filenum = FreeFile

If Dir(App.Path & "\Settings.ini") <> "" Then
    Kill App.Path & "\Settings.ini"
    filenum = FreeFile
End If

If closeApp = True Then
    If Dir(App.Path & "\Sch_Settings.ini") <> "" Then
        Kill App.Path & "\Sch_Settings.ini"
        filenum = FreeFile
    End If
End If



If Export_Map.chb_Scheduled.Value = 1 Then
    If closeApp = True Then
        Open App.Path & "\Sch_Settings.ini" For Append As #filenum
            Write #filenum, outputFormat
            Write #filenum, embedFonts
            Write #filenum, convMarkPoly
            Write #filenum, mapCompress
            Write #filenum, compType
            Write #filenum, pictSymb
            Write #filenum, colorMode
            Write #filenum, traspCol
            Write #filenum, layersAttrib
            Write #filenum, expGeoInfo
            Write #filenum, imageQuality
            Write #filenum, Export_Map.txb_dpi.Text
            Write #filenum, Export_Map.cmb_ExpFormat.ListIndex
            Write #filenum, Export_Map.cmb_ImageComp.ListIndex
            Write #filenum, Export_Map.cmb_PictSymb.ListIndex
            Write #filenum, Export_Map.cmb_ColorMode.ListIndex
            Write #filenum, Export_Map.cmb_LayersAttrib.ListIndex
            Write #filenum, Export_Map.chb_EmbedFonts.Value
            Write #filenum, Export_Map.chb_ConvMark.Value
            Write #filenum, Export_Map.chb_Compress.Value
            Write #filenum, Export_Map.chb_Progressive.Value
            Write #filenum, Export_Map.chb_ExpMapGeoInfo.Value
            'Write #filenum, color
            Write #filenum, Export_Map.sld_ImageQ.Value
    
            Dim dateStr As String
            dateStr = Format(Export_Map.DTPicker1.Value, "dd/mm/yyyy")
            Write #filenum, dateStr
            Write #filenum, Export_Map.txb_Time2Run.Text
            Write #filenum, Export_Map.opt_schOnce.Value
            Write #filenum, Export_Map.opt_schDaily.Value
            Write #filenum, Export_Map.opt_schCustom.Value
            Write #filenum, Export_Map.txb_Custom.Text
            Write #filenum, Export_Map.chb_Scheduled.Value
            Write #filenum, Export_Map.txb_ProgDir.Text
            Write #filenum, Export_Map.txb_OutDir.Text
            Write #filenum, Export_Map.chb_OutFolder.Value
            Write #filenum, Export_Map.chb_OutFolder.Enabled
            Write #filenum, Export_Map.chb_IncludeSub.Value
            Write #filenum, Export_Map.optMxd.Value
            Write #filenum, Export_Map.optMxt.Value
            Write #filenum, Export_Map.opt_ExpAll.Value
            Write #filenum, Export_Map.opt_ExpSel.Value
            Write #filenum, Export_Map.Frame4.Caption
        Close #filenum
        Exit Sub
    End If
End If


Open App.Path & "\Settings.ini" For Append As #filenum
    Write #filenum, outputFormat
    Write #filenum, embedFonts
    Write #filenum, convMarkPoly
    Write #filenum, mapCompress
    Write #filenum, compType
    Write #filenum, pictSymb
    Write #filenum, colorMode
    Write #filenum, traspCol
    Write #filenum, layersAttrib
    Write #filenum, expGeoInfo
    Write #filenum, imageQuality
    Write #filenum, Export_Map.txb_dpi.Text
    Write #filenum, Export_Map.cmb_ExpFormat.ListIndex
    Write #filenum, Export_Map.cmb_ImageComp.ListIndex
    Write #filenum, Export_Map.cmb_PictSymb.ListIndex
    Write #filenum, Export_Map.cmb_ColorMode.ListIndex
    Write #filenum, Export_Map.cmb_LayersAttrib.ListIndex
    Write #filenum, Export_Map.chb_EmbedFonts.Value
    Write #filenum, Export_Map.chb_ConvMark.Value
    Write #filenum, Export_Map.chb_Compress.Value
    Write #filenum, Export_Map.chb_Progressive.Value
    Write #filenum, Export_Map.chb_ExpMapGeoInfo.Value
    'Write #filenum, color
    Write #filenum, Export_Map.sld_ImageQ.Value
Close #filenum

Export_Map.cmd_SetAsDefault.Enabled = False


End Sub


Public Sub DefaultSettings()


Export_Map.txb_Info.Text = "With this tool is possible to export in batch ArcGIS projects maps " & _
                            "in six different output formats, at the desired resolution." & vbNewLine & vbNewLine & _
                            "Exported maps will be saved with the same name of the ArcGis projects, in the same directory, unless you specify a different output folder."

'initial form size
Export_Map.cmd_Settings.Caption = "- OPEN -"
Export_Map.Height = 5370


'general settings
extAGP = "mxd"
outputFormat = "jpg"
Export_Map.cmb_ExpFormat.Clear
Export_Map.cmb_ExpFormat.AddItem "EPS"
Export_Map.cmb_ExpFormat.AddItem "JPEG"
Export_Map.cmb_ExpFormat.AddItem "PDF"
Export_Map.cmb_ExpFormat.AddItem "PNG"
Export_Map.cmb_ExpFormat.AddItem "SVG"
Export_Map.cmb_ExpFormat.AddItem "TIF"
Export_Map.cmb_ExpFormat.ListIndex = 1
Export_Map.Frame4.Caption = "Projects to export: ALL (0)"
Export_Map.cmd_SelProjects.Enabled = False
Export_Map.chb_OutFolder.Enabled = False
Export_Map.chb_OutFolder.Value = 0
Export_Map.cmd_OutBrowse.Enabled = False
Export_Map.txb_OutDir.Enabled = False
Export_Map.txb_OutDir.BackColor = RGB(217, 217, 255)
exportError = False

includeSub = False
ReDim agpSelFileArray(0)

numOutput = 0
ReDim outputMaps(numOutput)


'output maps settings
Export_Map.chb_ConvMark.Visible = False
Export_Map.chb_EmbedFonts.Visible = False
Export_Map.chb_Compress.Visible = False
Export_Map.chb_Progressive.Visible = True
Export_Map.chb_Scheduled.Enabled = False
Export_Map.chb_ExpMapGeoInfo.Visible = False
Export_Map.opt_ExpAll.FontBold = True
Export_Map.opt_ExpAll.ForeColor = RGB(218, 27, 8)
Export_Map.optMxd.FontBold = True
Export_Map.optMxd.ForeColor = RGB(218, 27, 8)
Export_Map.cmb_PictSymb.Visible = False
Export_Map.cmb_ImageComp.Visible = False
Export_Map.cmb_LayersAttrib.Visible = False
Export_Map.Label17.Visible = False
Export_Map.Label4.Visible = False
Export_Map.Label3.Visible = False
Export_Map.Label7.Visible = False
Export_Map.Label8.Visible = False
Export_Map.cmd_Transparency.Visible = False
Export_Map.cmb_ColorMode.Clear
Export_Map.cmb_ColorMode.AddItem "24-bit True Color"
Export_Map.cmb_ColorMode.AddItem "8-bit grayscale"
Export_Map.cmb_ColorMode.ListIndex = 0
Export_Map.Label8.Caption = "No color"
Export_Map.Label9.Caption = "JPEG Quality"
Export_Map.cmb_PictSymb.Clear
Export_Map.cmb_PictSymb.AddItem "Rasterize layers with bitmap markers/fill"
Export_Map.cmb_PictSymb.AddItem "Rasterize layers with any picture marker/fill"
Export_Map.cmb_PictSymb.AddItem "Vectorize layers with bitmap markers/fill"
Export_Map.cmb_PictSymb.Text = "Rasterize layers with bitmap markers/fill"
If Export_Map.cmb_ExpFormat.ListIndex = 5 Then
    compType = esriTIFFCompressionDeflate
Else
    compType = esriExportImageCompressionDeflate
End If
mapCompress = False
convMarkPoly = False
embedFonts = False
expGeoInfo = False
pictSymb = esriPSORasterize
colorMode = esriExportImageTypeTrueColor
traspCol = False
layersAttrib = esriExportPDFLayerOptionsLayersOnly

Export_Map.txb_dpi.Text = ""

'scheduling settings
Export_Map.DTPicker1.Value = Format(Date, "dd/mm/yyyy")
Export_Map.opt_schOnce.FontBold = True
Export_Map.opt_schOnce.ForeColor = RGB(218, 27, 8)
Export_Map.Label14.Visible = False
Export_Map.Label15.Visible = False
Export_Map.txb_Custom.Visible = False

End Sub


Public Sub LoadSavedSettings()

On Error Resume Next

Dim iniExpFormat As Integer
Dim iniImageComp As Integer
Dim iniPyctSymb As Integer
Dim iniColorMode As Integer
Dim iniEmbedFonts As Integer
Dim iniConvMark As Integer
Dim iniCompress As Integer
Dim iniProgressive As Integer
Dim iniImageQ As Integer
Dim iniDPI As String
Dim iniSchDate As String
Dim iniSchTime As String
Dim iniSchOptOnce As Boolean
Dim iniSchOptDaily As Boolean
Dim iniSchOptCustom As Boolean
Dim iniSchInterval As String
Dim iniSchOnOff As Integer
Dim iniProgDir As String
Dim iniOutDir As String
Dim iniChbOutDir As Integer
Dim iniChbOutDirEnabled As Integer
Dim iniChbIncludeSub As Integer
Dim iniOptMxd As Boolean
Dim iniOptMxt As Boolean
Dim iniOptAll As Boolean
Dim iniOptSel As Boolean
Dim iniLabelFrame4 As String
Dim iniExpGeoInfo As String
Dim iniLayersAttrib As String



Export_Map.txb_Info.Text = "With this tool is possible to export in batch ArcGIS projects maps " & _
                            "in six different output formats, at the desired resolution." & vbNewLine & vbNewLine & _
                            "Exported maps will be saved with the same name of the ArcGis projects, in the same directory, unless you specify a different output folder."

'initial form size
Export_Map.cmd_Settings.Caption = "- OPEN -"
Export_Map.Height = 5370


Dim filenum As Integer
filenum = FreeFile


If closeApp = True Then
    Open App.Path & "\Sch_Settings.ini" For Input As #filenum
    Do While Not EOF(filenum)
        Input #filenum, outputFormat, embedFonts, convMarkPoly, mapCompress, compType, pictSymb, _
                        colorMode, traspCol, layersAttrib, expGeoInfo, imageQuality, iniDPI, iniExpFormat, iniImageComp, _
                        iniPyctSymb, iniColorMode, iniLayersAttrib, iniEmbedFonts, iniConvMark, iniCompress, _
                        iniProgressive, iniExpGeoInfo, iniImageQ, iniSchDate, iniSchTime, iniSchOptOnce, _
                        iniSchOptDaily, iniSchOptCustom, iniSchInterval, iniSchOnOff, iniProgDir, _
                        iniOutDir, iniChbOutDir, iniChbOutDirEnabled, iniChbIncludeSub, iniOptMxd, _
                        iniOptMxt, iniOptAll, iniOptSel, iniLabelFrame4, iniExpGeoInfo, iniLayersAttrib
    Loop
    Close #filenum
    Export_Map.Frame1.Enabled = False
    Export_Map.Frame2.Enabled = False
    Export_Map.Frame3.Enabled = False
    Export_Map.Frame4.Enabled = False
    Export_Map.cmd_Execute.Enabled = False
Else
    Open App.Path & "\Settings.ini" For Input As #filenum
    Do While Not EOF(filenum)
        Input #filenum, outputFormat, embedFonts, convMarkPoly, mapCompress, compType, pictSymb, _
                        colorMode, traspCol, layersAttrib, expGeoInfo, imageQuality, iniDPI, iniExpFormat, iniImageComp, _
                        iniPyctSymb, iniColorMode, iniLayersAttrib, iniEmbedFonts, iniConvMark, iniCompress, _
                        iniProgressive, iniExpGeoInfo, iniImageQ
    Loop
    Close #filenum
End If



'general settings
Export_Map.cmb_ExpFormat.Clear
Export_Map.cmb_ExpFormat.AddItem "EPS"
Export_Map.cmb_ExpFormat.AddItem "JPEG"
Export_Map.cmb_ExpFormat.AddItem "PDF"
Export_Map.cmb_ExpFormat.AddItem "PNG"
Export_Map.cmb_ExpFormat.AddItem "SVG"
Export_Map.cmb_ExpFormat.AddItem "TIF"
Export_Map.cmb_PictSymb.Clear
Export_Map.cmb_PictSymb.AddItem "Rasterize layers with bitmap markers/fill"
Export_Map.cmb_PictSymb.AddItem "Rasterize layers with any picture marker/fill"
Export_Map.cmb_PictSymb.AddItem "Vectorize layers with bitmap markers/fill"

    
If closeApp = False Then
    extAGP = "mxd"
    Export_Map.cmb_ExpFormat.Text = Export_Map.cmb_ExpFormat.Text
    Export_Map.Frame4.Caption = "Projects to export: ALL (0)"
    Export_Map.cmd_SelProjects.Enabled = False
    Export_Map.chb_OutFolder.Enabled = False
    Export_Map.chb_OutFolder.Value = 0
    Export_Map.cmd_OutBrowse.Enabled = False
    Export_Map.txb_OutDir.Enabled = False
    Export_Map.txb_OutDir.BackColor = RGB(217, 217, 255)
    Export_Map.cmb_PictSymb.Text = "Rasterize layers with bitmap markers/fill"
    Export_Map.opt_ExpAll.FontBold = True
    Export_Map.opt_ExpAll.ForeColor = RGB(218, 27, 8)
    Export_Map.optMxd.FontBold = True
    Export_Map.optMxd.ForeColor = RGB(218, 27, 8)
End If


exportError = False

Export_Map.txb_dpi.Text = iniDPI
dpi = iniDPI

Export_Map.cmb_ExpFormat.ListIndex = iniExpFormat


If Export_Map.cmb_ExpFormat.ListIndex = 0 Then         'EPS
    Export_Map.chb_ConvMark.Visible = True
    Export_Map.chb_EmbedFonts.Visible = True
    Export_Map.chb_Compress.Visible = False
    Export_Map.chb_Progressive.Visible = False
    Export_Map.chb_ExpMapGeoInfo.Visible = False
    Export_Map.cmb_PictSymb.Visible = True
    Export_Map.cmb_ImageComp.Visible = True
    Export_Map.cmb_ColorMode.Visible = True
    Export_Map.cmb_PictSymb.ListIndex = iniPyctSymb
    Export_Map.chb_ConvMark.Value = iniConvMark
    Export_Map.chb_EmbedFonts.Value = iniEmbedFonts
    Export_Map.cmb_PictSymb.Visible = True
    Export_Map.cmb_ImageComp.Visible = True
    Export_Map.cmb_ColorMode.Visible = True
    Export_Map.cmb_LayersAttrib.Visible = False
    Export_Map.Label4.Visible = True
    Export_Map.Label3.Visible = True
    Export_Map.Label5.Visible = True
    Export_Map.Label5.Caption = "Destination Colorspace:"
    Export_Map.cmb_ColorMode.Clear
    Export_Map.cmb_ColorMode.AddItem "RGB"
    Export_Map.cmb_ColorMode.AddItem "CMYK"
    Export_Map.cmb_ColorMode.ListIndex = iniColorMode
    colorMode = Export_Map.cmb_ColorMode.Text
    Export_Map.Label7.Visible = False
    Export_Map.Label8.Visible = False
    Export_Map.Label9.Visible = True
    Export_Map.Label9.Caption = "Output Image Quality"
    Export_Map.Label10.Visible = True
    Export_Map.Label17.Visible = False
    Export_Map.cmd_Transparency.Visible = False
    Export_Map.sld_ImageQ.Min = 1
    Export_Map.sld_ImageQ.Max = 5
    Export_Map.sld_ImageQ.SelStart = iniImageQ
    Export_Map.sld_ImageQ.Visible = True
    Export_Map.cmb_ImageComp.Clear
    Export_Map.cmb_ImageComp.AddItem "None"
    Export_Map.cmb_ImageComp.AddItem "RLE"
    Export_Map.cmb_ImageComp.AddItem "LZW"
    Export_Map.cmb_ImageComp.AddItem "Deflate"
    Export_Map.cmb_ImageComp.ListIndex = iniImageComp
    outputFormat = "eps"
ElseIf Export_Map.cmb_ExpFormat.ListIndex = 1 Then     'JPEG
    Export_Map.chb_ConvMark.Visible = False
    Export_Map.chb_EmbedFonts.Visible = False
    Export_Map.chb_Compress.Visible = False
    Export_Map.chb_Progressive.Visible = True
    Export_Map.chb_ExpMapGeoInfo.Visible = False
    Export_Map.cmb_PictSymb.Visible = False
    Export_Map.cmb_ImageComp.Visible = False
    Export_Map.cmb_ColorMode.Visible = True
    Export_Map.chb_Progressive.Caption = "Progressive"
    Export_Map.chb_Progressive.Value = iniProgressive
    Export_Map.Label4.Visible = False
    Export_Map.Label3.Visible = False
    Export_Map.Label5.Visible = True
    Export_Map.Label5.Caption = "Color mode:"
    Export_Map.cmb_ColorMode.Clear
    Export_Map.cmb_ColorMode.AddItem "24-bit True Color"
    Export_Map.cmb_ColorMode.AddItem "8-bit grayscale"
    Export_Map.cmb_ColorMode.ListIndex = iniColorMode
    Export_Map.cmb_LayersAttrib.Visible = False
    Export_Map.Label7.Visible = False
    Export_Map.Label8.Visible = False
    Export_Map.Label9.Visible = True
    Export_Map.Label9.Caption = "JPEG Quality"
    Export_Map.Label10.Visible = True
    Export_Map.Label17.Visible = False
    Export_Map.cmd_Transparency.Visible = False
    Export_Map.sld_ImageQ.Min = 0
    Export_Map.sld_ImageQ.Max = 100
    Export_Map.sld_ImageQ.SelStart = iniImageQ
    Export_Map.sld_ImageQ.Visible = True
    outputFormat = "jpg"
ElseIf Export_Map.cmb_ExpFormat.ListIndex = 2 Then     'PDF
    Export_Map.chb_ConvMark.Visible = True
    Export_Map.chb_EmbedFonts.Visible = True
    Export_Map.chb_Compress.Visible = True
    Export_Map.chb_Progressive.Visible = False
    Export_Map.chb_ExpMapGeoInfo.Visible = True
    Export_Map.cmb_PictSymb.Visible = True
    Export_Map.cmb_ImageComp.Visible = True
    Export_Map.cmb_ColorMode.Visible = True
    Export_Map.chb_Compress.Caption = "Compress vector/text graphics"
    Export_Map.chb_ConvMark.Value = iniConvMark
    Export_Map.chb_EmbedFonts.Value = iniEmbedFonts
    Export_Map.chb_Compress.Value = iniCompress
    Export_Map.chb_ExpMapGeoInfo.Value = expGeoInfo
    Export_Map.cmb_PictSymb.ListIndex = iniPyctSymb
    Export_Map.Label4.Visible = True
    Export_Map.Label3.Visible = True
    Export_Map.Label5.Visible = True
    Export_Map.Label5.Caption = "Destination Colorspace:"
    Export_Map.cmb_ColorMode.Clear
    Export_Map.cmb_ColorMode.AddItem "RGB"
    Export_Map.cmb_ColorMode.AddItem "CMYK"
    Export_Map.cmb_ColorMode.ListIndex = iniColorMode
    colorMode = Export_Map.cmb_ColorMode.Text
    Export_Map.cmb_LayersAttrib.Visible = True
    Export_Map.Label7.Visible = False
    Export_Map.Label8.Visible = False
    Export_Map.Label9.Visible = True
    Export_Map.Label9.Caption = "Output Image Quality"
    Export_Map.Label10.Visible = True
    Export_Map.Label17.Visible = True
    Export_Map.cmd_Transparency.Visible = False
    Export_Map.sld_ImageQ.Min = 1
    Export_Map.sld_ImageQ.Max = 5
    'Export_Map.sld_ImageQ.SelStart = iniImageQ
    Export_Map.sld_ImageQ.Value = iniImageQ
    Export_Map.sld_ImageQ.Visible = True
    Export_Map.cmb_ImageComp.Clear
    Export_Map.cmb_ImageComp.AddItem "None"
    Export_Map.cmb_ImageComp.AddItem "RLE"
    Export_Map.cmb_ImageComp.AddItem "LZW"
    Export_Map.cmb_ImageComp.AddItem "Deflate"
    Export_Map.cmb_ImageComp.ListIndex = iniImageComp
    Export_Map.cmb_LayersAttrib.Clear
    Export_Map.cmb_LayersAttrib.AddItem "None"
    Export_Map.cmb_LayersAttrib.AddItem "Export PDF Layers Only"
    Export_Map.cmb_LayersAttrib.AddItem "Export PDF Layers and Feature Attributes"
    Export_Map.cmb_LayersAttrib.ListIndex = iniLayersAttrib
    outputFormat = "pdf"
    Export_Map.chb_Compress.ToolTipText = "For PDF it compresses only vector and text portions of the map"
ElseIf Export_Map.cmb_ExpFormat.ListIndex = 3 Then     'PNG
    Export_Map.chb_ConvMark.Visible = False
    Export_Map.chb_EmbedFonts.Visible = False
    Export_Map.chb_Compress.Visible = False
    Export_Map.chb_Progressive.Visible = True
    Export_Map.chb_ExpMapGeoInfo.Visible = False
    Export_Map.cmb_PictSymb.Visible = False
    Export_Map.cmb_ImageComp.Visible = False
    Export_Map.cmb_ColorMode.Visible = True
    Export_Map.cmb_LayersAttrib.Visible = False
    Export_Map.chb_Progressive.Caption = "Interlaced"
    Export_Map.chb_Progressive.Value = iniProgressive
    Export_Map.Label4.Visible = False
    Export_Map.Label3.Visible = False
    Export_Map.Label5.Visible = True
    Export_Map.Label5.Caption = "Color mode:"
    Export_Map.cmb_ColorMode.Clear
    Export_Map.cmb_ColorMode.AddItem "24-bit True Color"
    Export_Map.cmb_ColorMode.AddItem "8-bit grayscale"
    Export_Map.cmb_ColorMode.ListIndex = iniColorMode
    Export_Map.Label7.Visible = True
    Export_Map.Label8.Visible = True
    Export_Map.Label9.Visible = False
    Export_Map.Label10.Visible = False
    Export_Map.Label17.Visible = False
    Export_Map.sld_ImageQ.Visible = False
    Export_Map.cmd_Transparency.Visible = True
    outputFormat = "png"
ElseIf Export_Map.cmb_ExpFormat.ListIndex = 4 Then     'SVG
    Export_Map.chb_ConvMark.Visible = True
    Export_Map.chb_EmbedFonts.Visible = True
    Export_Map.chb_Compress.Visible = True
    Export_Map.chb_Progressive.Visible = False
    Export_Map.chb_ExpMapGeoInfo.Visible = False
    Export_Map.cmb_PictSymb.Visible = True
    Export_Map.cmb_ImageComp.Visible = False
    Export_Map.cmb_ColorMode.Visible = False
    Export_Map.cmb_LayersAttrib.Visible = False
    Export_Map.chb_Compress.Caption = "Compress document"
    Export_Map.chb_ConvMark.Value = iniConvMark
    Export_Map.chb_EmbedFonts.Value = iniEmbedFonts
    Export_Map.chb_Compress.Value = iniCompress
    Export_Map.cmb_PictSymb.ListIndex = iniPyctSymb
    Export_Map.Label7.Visible = False
    Export_Map.Label8.Visible = False
    Export_Map.Label9.Visible = True
    Export_Map.Label9.Caption = "Output Image Quality"
    Export_Map.Label10.Visible = True
    Export_Map.Label17.Visible = False
    Export_Map.cmd_Transparency.Visible = False
    Export_Map.sld_ImageQ.Min = 1
    Export_Map.sld_ImageQ.Max = 5
    Export_Map.sld_ImageQ.SelStart = iniImageQ
    Export_Map.sld_ImageQ.Visible = True
    Export_Map.Label4.Visible = False
    Export_Map.Label3.Visible = True
    Export_Map.Label5.Visible = False
    outputFormat = "svg"
    Export_Map.chb_Compress.ToolTipText = "For SVG changes the file extension to *.svgz"
ElseIf Export_Map.cmb_ExpFormat.ListIndex = 5 Then     'TIF
    Export_Map.chb_ConvMark.Visible = False
    Export_Map.chb_EmbedFonts.Visible = False
    Export_Map.chb_Compress.Visible = False
    Export_Map.chb_Progressive.Visible = False
    Export_Map.chb_ExpMapGeoInfo.Visible = False
    Export_Map.cmb_PictSymb.Visible = False
    Export_Map.cmb_ImageComp.Visible = True
    Export_Map.cmb_ColorMode.Visible = True
    Export_Map.cmb_LayersAttrib.Visible = False
    Export_Map.Label4.Visible = True
    Export_Map.Label3.Visible = False
    Export_Map.Label5.Visible = True
    Export_Map.Label5.Caption = "Color mode:"
    Export_Map.cmb_ColorMode.Clear
    Export_Map.cmb_ColorMode.AddItem "24-bit True Color"
    Export_Map.cmb_ColorMode.AddItem "8-bit grayscale"
    Export_Map.cmb_ColorMode.ListIndex = iniColorMode
    Export_Map.Label7.Visible = False
    Export_Map.Label8.Visible = False
    Export_Map.Label9.Visible = True
    Export_Map.Label9.Caption = "Deflate Quality"
    Export_Map.Label10.Visible = True
    Export_Map.Label17.Visible = False
    Export_Map.cmd_Transparency.Visible = False
    Export_Map.sld_ImageQ.Min = 0
    Export_Map.sld_ImageQ.Max = 100
    Export_Map.sld_ImageQ.SelStart = iniImageQ
    Export_Map.sld_ImageQ.Visible = True
    Export_Map.sld_ImageQ.Enabled = True
    Export_Map.cmb_ImageComp.Clear
    Export_Map.cmb_ImageComp.AddItem "None"
    Export_Map.cmb_ImageComp.AddItem "LZW"
    Export_Map.cmb_ImageComp.AddItem "Deflate"
    Export_Map.cmb_ImageComp.AddItem "Pack Bits"
    Export_Map.cmb_ImageComp.AddItem "JPEG"
    Export_Map.cmb_ImageComp.ListIndex = iniImageComp
    outputFormat = "tif"
End If


If outputFormat = "svg" Or outputFormat = "pdf" Or outputFormat = "eps" Then
    If iniImageQ = 1 Then
        imageQuality = esriRasterOutputDraft
    ElseIf iniImageQ = 2 Then
        imageQuality = esriRasterOutputDraft
    ElseIf iniImageQ = 3 Then
        imageQuality = esriRasterOutputNormal
    ElseIf iniImageQ = 4 Then
        imageQuality = esriRasterOutputBest
    ElseIf iniImageQ = 5 Then
        imageQuality = esriRasterOutputBest
    End If
End If


'other settings
If closeApp = True Then
    Export_Map.frame_Options.Enabled = False
    Export_Map.frame_Schedule.Enabled = False
    Export_Map.DTPicker1.Value = Format(iniSchDate, "dd/mm/yyyy")
    Export_Map.txb_Time2Run.Text = iniSchTime
    Export_Map.opt_schOnce.Value = iniSchOptOnce
    Export_Map.opt_schDaily.Value = iniSchOptDaily
    Export_Map.opt_schCustom.Value = iniSchOptCustom
    If iniSchOptOnce = True Then
        Export_Map.opt_schOnce.FontBold = True
        Export_Map.opt_schOnce.ForeColor = RGB(218, 27, 8)
    ElseIf iniSchOptDaily = True Then
        Export_Map.opt_schDaily.FontBold = True
        Export_Map.opt_schDaily.ForeColor = RGB(218, 27, 8)
    ElseIf iniSchOptCustom = True Then
        Export_Map.opt_schCustom.FontBold = True
        Export_Map.opt_schCustom.ForeColor = RGB(218, 27, 8)
    End If
    If iniSchOptCustom = True Then
        Export_Map.Label14.Visible = True
        Export_Map.Label15.Visible = True
        Export_Map.txb_Custom.Visible = True
        Export_Map.txb_Custom.Text = iniSchInterval
    Else
        Export_Map.Label14.Visible = False
        Export_Map.Label15.Visible = False
        Export_Map.txb_Custom.Visible = False
    End If
    Export_Map.chb_Scheduled.Value = iniSchOnOff
    Export_Map.chb_Scheduled.FontBold = True
    Export_Map.chb_Scheduled.ForeColor = RGB(218, 27, 8)
    Export_Map.txb_ProgDir.Text = iniProgDir
    Export_Map.txb_OutDir.Text = iniOutDir
    Export_Map.chb_OutFolder.Value = iniChbOutDir
    Export_Map.chb_OutFolder.Enabled = iniChbOutDirEnabled
    If iniChbOutDirEnabled = 0 Then
        Export_Map.txb_OutDir.BackColor = RGB(217, 217, 255)
    Else
        Export_Map.txb_OutDir.BackColor = RGB(255, 255, 255)
    End If
    Export_Map.chb_IncludeSub.Value = iniChbIncludeSub
    If iniChbIncludeSub = 1 Then
        Export_Map.chb_IncludeSub.FontBold = True
        Export_Map.chb_IncludeSub.ForeColor = RGB(218, 27, 8)
    End If
    Export_Map.optMxd.Value = iniOptMxd
    Export_Map.optMxt.Value = iniOptMxt
    If iniOptMxd = True Then
        Export_Map.optMxd.FontBold = True
        Export_Map.optMxd.ForeColor = RGB(218, 27, 8)
    ElseIf iniOptMxt = True Then
        Export_Map.optMxt.FontBold = True
        Export_Map.optMxt.ForeColor = RGB(218, 27, 8)
    End If
    Export_Map.opt_ExpAll.Value = iniOptAll
    Export_Map.opt_ExpSel.Value = iniOptSel
    If iniOptAll = True Then
        Export_Map.opt_ExpAll.FontBold = True
        Export_Map.opt_ExpAll.ForeColor = RGB(218, 27, 8)
    ElseIf iniOptSel = True Then
        Export_Map.opt_ExpSel.FontBold = True
        Export_Map.opt_ExpSel.ForeColor = RGB(218, 27, 8)
    End If
    Export_Map.Frame4.Caption = iniLabelFrame4
Else
    Export_Map.DTPicker1.Value = Format(Date, "dd/mm/yyyy")
    Export_Map.opt_schOnce.FontBold = True
    Export_Map.opt_schOnce.ForeColor = RGB(218, 27, 8)
    Export_Map.Label14.Visible = False
    Export_Map.Label15.Visible = False
    Export_Map.txb_Custom.Visible = False
End If

End Sub
