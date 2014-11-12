VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Export_Map 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Map Tool (1.7.3)"
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   Icon            =   "Export_map.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MousePointer    =   4  'Icon
   ScaleHeight     =   9510
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txb_Info 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Bitstream Vera Sans Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1260
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   105
      Width           =   7230
   End
   Begin VB.CommandButton cmd_RestoreDef 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   945
      MaskColor       =   &H8000000F&
      TabIndex        =   53
      ToolTipText     =   "Restore output parameters settings to the original default values"
      Top             =   4515
      Width           =   585
   End
   Begin VB.CommandButton cmd_SetAsDefault 
      Caption         =   "Save as Default"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   225
      MaskColor       =   &H8000000F&
      TabIndex        =   52
      ToolTipText     =   "Set current settings as default"
      Top             =   4200
      Width           =   1305
   End
   Begin VB.Frame frame_Schedule 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Schedule task"
      ForeColor       =   &H000040C0&
      Height          =   975
      Left            =   225
      TabIndex        =   41
      Top             =   8400
      Width           =   7230
      Begin VB.TextBox txb_Custom 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   5685
         TabIndex        =   55
         Top             =   600
         Width           =   570
      End
      Begin VB.OptionButton opt_schCustom 
         BackColor       =   &H00C0FFC0&
         Caption         =   "other"
         Height          =   195
         Left            =   4395
         TabIndex        =   51
         ToolTipText     =   "It runs every < custom > days at the same time"
         Top             =   675
         Width           =   765
      End
      Begin VB.OptionButton opt_schDaily 
         BackColor       =   &H00C0FFC0&
         Caption         =   "daily"
         Height          =   195
         Left            =   4395
         TabIndex        =   50
         ToolTipText     =   "It runs each day at the same time"
         Top             =   435
         Width           =   975
      End
      Begin VB.OptionButton opt_schOnce 
         BackColor       =   &H00C0FFC0&
         Caption         =   "once"
         Height          =   195
         Left            =   4395
         TabIndex        =   49
         ToolTipText     =   "It runs once at scheduled time"
         Top             =   180
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txb_Time2Run 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2970
         TabIndex        =   47
         ToolTipText     =   "Time format allowed: hh.mm.ss"
         Top             =   555
         Width           =   1305
      End
      Begin VB.CheckBox chb_Scheduled 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Scheduled"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5700
         TabIndex        =   44
         ToolTipText     =   "Set scheduled task to TRUE/FALSE"
         Top             =   180
         Width           =   1440
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   1275
         Top             =   75
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   2970
         TabIndex        =   48
         ToolTipText     =   "Date selector"
         Top             =   225
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   16761024
         Format          =   16449537
         CurrentDate     =   39520
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFC0&
         Caption         =   "days"
         Height          =   210
         Left            =   6270
         TabIndex        =   56
         Top             =   660
         Width           =   450
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFC0&
         Caption         =   "every"
         Height          =   210
         Left            =   5235
         TabIndex        =   54
         Top             =   660
         Width           =   450
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Date to run:"
         Height          =   240
         Left            =   2055
         TabIndex        =   46
         Top             =   315
         Width           =   885
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Time to run:"
         Height          =   240
         Left            =   2055
         TabIndex        =   45
         Top             =   615
         Width           =   885
      End
      Begin VB.Label lbl_CurrentTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   210
         TabIndex        =   43
         Top             =   510
         Width           =   1680
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Current time:"
         Height          =   240
         Left            =   195
         TabIndex        =   42
         Top             =   270
         Width           =   960
      End
   End
   Begin VB.CommandButton cmd_Settings 
      Caption         =   "- OPEN  -"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   225
      MaskColor       =   &H8000000F&
      TabIndex        =   40
      ToolTipText     =   "Expand/collapse form to set output parameters and to schedule task"
      Top             =   4515
      Width           =   705
   End
   Begin VB.CheckBox chb_OutFolder 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Output folder for maps:"
      Height          =   330
      Left            =   225
      TabIndex        =   39
      Top             =   2340
      Width           =   3030
   End
   Begin VB.TextBox txb_OutDir 
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   225
      TabIndex        =   38
      Top             =   2670
      Width           =   3855
   End
   Begin VB.CommandButton cmd_OutBrowse 
      Caption         =   "Browse..."
      Height          =   285
      Left            =   4185
      TabIndex        =   37
      Top             =   2670
      Width           =   900
   End
   Begin VB.CheckBox chb_IncludeSub 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Include subfolders"
      Height          =   285
      Left            =   5655
      TabIndex        =   33
      ToolTipText     =   "Search for ArcGIS projects in all subfolders"
      Top             =   1905
      Width           =   1950
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Projects to export: ALL"
      ForeColor       =   &H000040C0&
      Height          =   735
      Left            =   240
      TabIndex        =   29
      Top             =   3360
      Width           =   2535
      Begin VB.CommandButton cmd_SelProjects 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1965
         TabIndex        =   32
         ToolTipText     =   "Browse for projects selection"
         Top             =   300
         Width           =   420
      End
      Begin VB.OptionButton opt_ExpSel 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Selected"
         Height          =   255
         Left            =   840
         TabIndex        =   31
         Top             =   315
         Width           =   1125
      End
      Begin VB.OptionButton opt_ExpAll 
         BackColor       =   &H00C0FFC0&
         Caption         =   "All"
         Height          =   255
         Left            =   150
         TabIndex        =   30
         Top             =   315
         Value           =   -1  'True
         Width           =   555
      End
   End
   Begin VB.Frame frame_Options 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Output options"
      ForeColor       =   &H000040C0&
      Height          =   3225
      Left            =   240
      TabIndex        =   13
      Top             =   4980
      Width           =   7215
      Begin VB.CheckBox chb_ExpMapGeoInfo 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Export Map Georeference Information"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         ToolTipText     =   "Available to Adobe Acrobat and Adobe Reader vers. 9"
         Top             =   2775
         Width           =   3300
      End
      Begin VB.ComboBox cmb_LayersAttrib 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   2310
         Width           =   3975
      End
      Begin MSComctlLib.Slider sld_ImageQ 
         Height          =   285
         Left            =   5760
         TabIndex        =   34
         Top             =   2790
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Min             =   1
         Max             =   5
         SelStart        =   3
         TickStyle       =   3
         Value           =   3
      End
      Begin VB.CommandButton cmd_Transparency 
         Caption         =   "color"
         Height          =   300
         Left            =   5745
         TabIndex        =   26
         ToolTipText     =   "Browse for a color selector"
         Top             =   2205
         Width           =   810
      End
      Begin VB.CheckBox chb_Progressive 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Progressive"
         Height          =   255
         Left            =   4560
         TabIndex        =   23
         Top             =   1680
         Width           =   1320
      End
      Begin VB.ComboBox cmb_ColorMode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1680
         Width           =   3975
      End
      Begin VB.ComboBox cmb_ImageComp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   480
         Width           =   3975
      End
      Begin VB.CheckBox chb_Compress 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Compress vector/text graphics"
         Height          =   255
         Left            =   4560
         TabIndex        =   18
         Top             =   1290
         Width           =   2580
      End
      Begin VB.ComboBox cmb_PictSymb 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1080
         Width           =   3975
      End
      Begin VB.CheckBox chb_ConvMark 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Convert marker symbols to polygons"
         Height          =   375
         Left            =   4560
         TabIndex        =   15
         Top             =   780
         Width           =   2415
      End
      Begin VB.CheckBox chb_EmbedFonts 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Embed all documents fonts"
         Height          =   255
         Left            =   4560
         TabIndex        =   14
         Top             =   420
         Width           =   2415
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Layer and Attributes:"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   2070
         Width           =   2385
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "LOW                HIGH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5760
         TabIndex        =   36
         Top             =   2625
         Width           =   1290
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Caption         =   "Output Image Quality:"
         Height          =   255
         Left            =   4140
         TabIndex        =   35
         Top             =   2820
         Width           =   1560
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6615
         TabIndex        =   28
         Top             =   2235
         Width           =   420
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Set transparency color:"
         Height          =   270
         Left            =   4050
         TabIndex        =   27
         Top             =   2265
         Width           =   1680
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Color mode:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   2385
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Image compression:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Picture symbol:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Output"
      ForeColor       =   &H000040C0&
      Height          =   735
      Left            =   4920
      TabIndex        =   12
      Top             =   3360
      Width           =   2535
      Begin VB.ComboBox cmb_ExpFormat 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Export format:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Resolution (96 to 1200)"
      ForeColor       =   &H000040C0&
      Height          =   735
      Left            =   2880
      TabIndex        =   8
      Top             =   3360
      Width           =   1935
      Begin VB.TextBox txb_dpi 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   480
         TabIndex        =   9
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "dpi:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   330
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ArcGis project"
      ForeColor       =   &H000040C0&
      Height          =   690
      Left            =   5640
      TabIndex        =   4
      Top             =   2415
      Width           =   1815
      Begin VB.OptionButton optMxt 
         BackColor       =   &H00C0FFC0&
         Caption         =   "MXT"
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optMxd 
         BackColor       =   &H00C0FFC0&
         Caption         =   "MXD"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.CommandButton cmd_Execute 
      Caption         =   "EXECUTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6150
      TabIndex        =   3
      Top             =   4425
      Width           =   1305
   End
   Begin VB.CommandButton cmd_InBrowse 
      Caption         =   "Browse..."
      Height          =   285
      Left            =   4200
      TabIndex        =   0
      Top             =   1905
      Width           =   885
   End
   Begin VB.TextBox txb_ProgDir 
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   1905
      Width           =   3855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Original source code by Alberto Laurenti, Rome (Italy)"
      ForeColor       =   &H00004000&
      Height          =   225
      Left            =   3615
      MouseIcon       =   "Export_map.frx":17D2A
      MousePointer    =   99  'Custom
      TabIndex        =   57
      ToolTipText     =   "mailto: a_laurenti@hotmail.com"
      Top             =   1350
      Width           =   3855
   End
   Begin VB.Label lbl_status 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1695
      TabIndex        =   7
      Top             =   4470
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ArcGis project input folder:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1665
      Width           =   2790
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Export_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================================================================
'======================================================================
'EXPORTMAP.EXE
'
'Author:    Alberto Laurenti (Agriconsulting S.p.A. - Rome, Italy)
'Date:      29 january 2008
'Purpose:   program to export ArcGis maps in batch.
'Contact:   a_laurenti@hotmail.com
'======================================================================
'======================================================================


'NOTE: be sure that each ArcGIS project has been saved in layout view.


Option Explicit

Dim Result As Long
Dim getdir As String
Dim fileAGP As String
Dim totAGPFiles As Integer
Dim dirAGP As String
Dim progMode As Boolean
Dim pCSS As IExportColorspaceSettings
Dim pOutColor As IColor
Dim progDate As Date
Dim suffixReport As String
Dim customDays As Integer
Dim running As Boolean


Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private m_pAoInitialize As IAoInitialize



Private Sub chb_ExpMapGeoInfo_Click()

If chb_ExpMapGeoInfo.Value = 1 Then
    expGeoInfo = 1
Else
    expGeoInfo = 0
End If

End Sub

Private Sub chb_IncludeSub_Click()

If chb_IncludeSub.Value = 1 Then
    If opt_ExpSel.Value = True Then
        MsgBox "You can't include subfolders with 'Selected' projects to export option checked!", vbExclamation, "Warning"
        chb_IncludeSub.Value = 0
        opt_ExpAll.Value = True
        includeSub = False
        Exit Sub
    Else
        chb_IncludeSub.FontBold = True
        chb_IncludeSub.ForeColor = RGB(218, 27, 8)
        If txb_ProgDir.Text <> "" Then
            '-------------------------
            Call LoopFolder.LoopFolderList
            '-------------------------
            If totAGP > 0 Then
                If agpSubFolders = False Then
                    MsgBox "No subfolders found!", vbInformation, "Path to projects"
                    chb_IncludeSub.Value = 0
                    chb_IncludeSub.FontBold = False
                    chb_IncludeSub.ForeColor = RGB(0, 0, 0)
                    includeSub = False
                Else
                    MsgBox "You're going to export  << " & totAGP & " >>  ArcGis " & UCase(extAGP) & " projects.", vbInformation, "Projects to export"
                    Frame4.Caption = "Projects to export: ALL (" & totAGP & ")"
                End If
            Else
                Frame4.Caption = "Projects to export: ALL (0)"
            End If
        End If
        includeSub = True
    End If
Else
    If txb_ProgDir.Text <> "" Then
        '-------------------------
        Call LoopFolder.LoopFolderList
        '-------------------------
        If totAGP > 0 Then
            MsgBox "You're going to export  << " & totAGP & " >>  ArcGis " & UCase(extAGP) & " projects.", vbInformation, "Projects to export"
            Frame4.Caption = "Projects to export: ALL (" & totAGP & ")"
        Else
            Frame4.Caption = "Projects to export: ALL (0)"
        End If
    End If
    chb_IncludeSub.FontBold = False
    chb_IncludeSub.ForeColor = RGB(0, 0, 0)
    includeSub = False
End If

End Sub


Private Sub chb_OutFolder_Click()

If chb_OutFolder.Value = 1 Then
    cmd_OutBrowse.Enabled = True
    chb_OutFolder.FontBold = True
    chb_OutFolder.ForeColor = RGB(218, 27, 8)
    txb_OutDir.BackColor = RGB(255, 255, 255)
    txb_OutDir.Enabled = True
Else
    cmd_OutBrowse.Enabled = False
    chb_OutFolder.FontBold = False
    chb_OutFolder.ForeColor = RGB(0, 0, 0)
    txb_OutDir.BackColor = RGB(217, 217, 255)
    txb_OutDir.Enabled = False
End If

End Sub


Private Sub chb_Scheduled_Click()

If chb_Scheduled.Value = 1 Then
    If txb_Time2Run.Text = Format(txb_Time2Run.Text, "hh.mm.ss") Then
        chb_Scheduled.FontBold = True
        chb_Scheduled.ForeColor = RGB(218, 27, 8)
    Else
        chb_Scheduled.Value = 0
        MsgBox "ATTENTION: time format must be 'hh.mm.ss'!", vbExclamation, "Warning"
    End If
    If opt_schCustom.Value = True And txb_Custom.Text = "" Then
        chb_Scheduled.Value = 0
        MsgBox "ATTENTION: you have to specify the interval (days) for scheduled task!", vbExclamation, "Warning"
    End If
Else
    chb_Scheduled.FontBold = False
    chb_Scheduled.ForeColor = RGB(0, 0, 0)
End If

End Sub


Private Sub cmb_LayersAttrib_Click()

If cmb_LayersAttrib.ListIndex = 0 Then
    layersAttrib = esriExportPDFLayerOptionsNone
ElseIf cmb_LayersAttrib.ListIndex = 1 Then
    layersAttrib = esriExportPDFLayerOptionsLayersOnly
ElseIf cmb_LayersAttrib.ListIndex = 2 Then
    layersAttrib = esriExportPDFLayerOptionsLayersAndFeatureAttributes
End If

End Sub

Private Sub cmd_RestoreDef_Click()

'-------------------------------
Call Settings.DefaultSettings
'-------------------------------

If Dir(App.Path & "\Settings.ini") <> "" Then
    Kill App.Path & "\Settings.ini"
End If

End Sub

Private Sub cmd_SetAsDefault_Click()

'-------------------------------
Call Settings.SaveAsDefaults
'-------------------------------

End Sub


Private Sub Label16_Click()

Dim email As String

email = "mailto: a_laurenti@hotmail.com"

ShellExecute Me.hwnd, vbNullString, email, vbNullString, vbNullString, 1

End Sub

Private Sub opt_schCustom_Click()

If opt_schCustom.Value = True Then
    opt_schCustom.FontBold = True
    opt_schCustom.ForeColor = RGB(218, 27, 8)
    opt_schDaily.FontBold = False
    opt_schDaily.ForeColor = RGB(0, 0, 0)
    opt_schOnce.FontBold = False
    opt_schOnce.ForeColor = RGB(0, 0, 0)
    Label14.Visible = True
    Label15.Visible = True
    txb_Custom.Visible = True
End If

End Sub

Private Sub opt_schDaily_Click()

If opt_schDaily.Value = True Then
    opt_schDaily.FontBold = True
    opt_schDaily.ForeColor = RGB(218, 27, 8)
    opt_schOnce.FontBold = False
    opt_schOnce.ForeColor = RGB(0, 0, 0)
    opt_schCustom.FontBold = False
    opt_schCustom.ForeColor = RGB(0, 0, 0)
    Label14.Visible = False
    Label15.Visible = False
    txb_Custom.Visible = False
End If

End Sub

Private Sub opt_schOnce_Click()

If opt_schOnce.Value = True Then
    opt_schOnce.FontBold = True
    opt_schOnce.ForeColor = RGB(218, 27, 8)
    opt_schDaily.FontBold = False
    opt_schDaily.ForeColor = RGB(0, 0, 0)
    opt_schCustom.FontBold = False
    opt_schCustom.ForeColor = RGB(0, 0, 0)
    Label14.Visible = False
    Label15.Visible = False
    txb_Custom.Visible = False
End If

End Sub

Private Sub Timer1_Timer()

If lbl_CurrentTime.Caption <> CStr(Now) Then
    lbl_CurrentTime.Caption = Now
End If

End Sub


Private Sub cmd_InBrowse_Click()

getdir = BrowseForFolder(Me, "Select a folder for the ArcGis projects:", txb_ProgDir.Text)
If Len(getdir) = 0 Then Exit Sub  'user selected cancel
txb_ProgDir.Text = getdir
txb_OutDir.Text = txb_ProgDir.Text
chb_OutFolder.Enabled = True
opt_ExpAll.Value = True

'-------------------------
Call LoopFolder.LoopFolderList
'-------------------------

If totAGP > 0 Then
    MsgBox "You're going to export  << " & totAGP & " >>  ArcGis " & UCase(extAGP) & " projects.", vbInformation, "Projects to export"
    Frame4.Caption = "Projects to export: ALL (" & totAGP & ")"
Else
    Frame4.Caption = "Projects to export: ALL (0)"
End If

End Sub

Private Sub cmd_OutBrowse_Click()

getdir = BrowseForFolder(Me, "Select a folder for the output maps:", txb_ProgDir.Text)
If Len(getdir) = 0 Then Exit Sub  'user selected cancel
txb_OutDir.Text = getdir

End Sub

Private Sub cmd_Settings_Click()

If cmd_Settings.Caption = "- OPEN -" Then
    cmd_Settings.Caption = "- CLOSE -"
    Export_Map.Height = 9930
Else
    cmd_Settings.Caption = "- OPEN -"
    Export_Map.Height = 5370
End If

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'this procedure receives the callbacks from the System Tray icon.
Dim msg As Long

'the value of X will vary depending upon the scalemode setting
If Me.ScaleMode = vbPixels Then
    msg = X
Else
    msg = X / Screen.TwipsPerPixelX
End If

Select Case msg
    Case WM_LBUTTONUP        '514 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
    Case WM_LBUTTONDBLCLK    '515 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
    Case WM_RBUTTONUP        '517 display popup menu
            Result = SetForegroundWindow(Me.hwnd)
            Me.PopupMenu Me.mPopupSys
End Select

End Sub

Private Sub Form_Resize()

'this is necessary to assure that the minimized window is hidden
If Me.WindowState = vbMinimized Then Me.Hide


End Sub

Private Sub Form_Unload(Cancel As Integer)


If chb_Scheduled.Value = 1 Then

    Dim answ2 As Integer
    answ2 = MsgBox("Tool is running. Are you sure you want to exit the application?", vbYesNo, "Close")
    If answ2 = 6 Then
        running = False
        'Shutdown
        m_pAoInitialize.Shutdown
        'this removes the icon from the system tray
        Shell_NotifyIcon NIM_DELETE, nid
        Exit Sub
        Unload Me
    Else
        closeApp = True
        '-----------------------------
        Call Settings.SaveAsDefaults
        '-----------------------------
        Exit Sub
    End If

Else

    'Shutdown
    m_pAoInitialize.Shutdown
    
    'this removes the icon from the system tray
    Shell_NotifyIcon NIM_DELETE, nid
        
    Unload Me

End If



End Sub

Private Function CheckOutLicenses(productCode As esriLicenseProductCode) As esriLicenseStatus
  
  Dim licenseStatus As esriLicenseStatus
  Set m_pAoInitialize = New AoInitialize
  CheckOutLicenses = esriLicenseUnavailable
    
  'Check the productCode
  licenseStatus = m_pAoInitialize.IsProductCodeAvailable(productCode)
  If (licenseStatus = esriLicenseAvailable) Then
      'Initialize the license
      licenseStatus = m_pAoInitialize.Initialize(productCode)
  End If
  
  CheckOutLicenses = licenseStatus
  
End Function



Private Sub chb_Compress_Click()

If chb_Compress.Value = 1 Then
    mapCompress = True
Else
    mapCompress = False
End If

cmd_SetAsDefault.Enabled = True

End Sub

Private Sub chb_ConvMark_Click()

If chb_ConvMark.Value = 1 Then
    convMarkPoly = True
Else
    convMarkPoly = False
End If

cmd_SetAsDefault.Enabled = True

End Sub

Private Sub chb_EmbedFonts_Click()

If chb_EmbedFonts.Value = 1 Then
    embedFonts = True
Else
    embedFonts = False
End If

cmd_SetAsDefault.Enabled = True

End Sub

Private Sub chb_Progressive_Click()

If chb_Progressive.Value = 1 Then
    progMode = True
Else
    progMode = False
End If

cmd_SetAsDefault.Enabled = True

End Sub


Private Sub cmb_ColorMode_LostFocus()

If cmb_ColorMode.ListIndex = 0 Then
    If cmb_ExpFormat.ListIndex = 1 Or cmb_ExpFormat.ListIndex = 5 Then      'JPEG - TIF
        colorMode = esriExportImageTypeTrueColor
    Else
        colorMode = "RGB"
    End If
ElseIf cmb_ColorMode.ListIndex = 1 Then
    If cmb_ExpFormat.ListIndex = 1 Or cmb_ExpFormat.ListIndex = 5 Then      'JPEG - TIF
        colorMode = esriExportImageTypeGrayscale
    Else
        colorMode = "CMYK"
    End If
End If

cmd_SetAsDefault.Enabled = True

End Sub


Private Sub cmb_ExpFormat_Click()

If cmb_ExpFormat.ListIndex = 0 Then         'EPS
    chb_ConvMark.Visible = True
    chb_EmbedFonts.Visible = True
    chb_Compress.Visible = False
    chb_Progressive.Visible = False
    chb_ExpMapGeoInfo.Visible = False
    chb_ConvMark.Value = 0
    chb_EmbedFonts.Value = 0
    cmb_PictSymb.Visible = True
    cmb_ImageComp.Visible = True
    cmb_ColorMode.Visible = True
    cmb_LayersAttrib.Visible = False
    Label4.Visible = True
    Label3.Visible = True
    Label5.Visible = True
    Label5.Caption = "Destination Colorspace:"
    cmb_ColorMode.Clear
    cmb_ColorMode.AddItem "RGB"
    cmb_ColorMode.AddItem "CMYK"
    cmb_ColorMode.Text = "RGB"
    colorMode = "RGB"
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = True
    Label9.Caption = "Output Image Quality"
    Label10.Visible = True
    Label17.Visible = False
    cmd_Transparency.Visible = False
    sld_ImageQ.Min = 1
    sld_ImageQ.Max = 5
    sld_ImageQ.SelStart = 3
    sld_ImageQ.Visible = True
    imageQuality = esriRasterOutputNormal
    cmb_ImageComp.Clear
    cmb_ImageComp.AddItem "None"
    cmb_ImageComp.AddItem "RLE"
    cmb_ImageComp.AddItem "LZW"
    cmb_ImageComp.AddItem "Deflate"
    cmb_ImageComp.Text = "Deflate"
    outputFormat = "eps"
ElseIf cmb_ExpFormat.ListIndex = 1 Then     'JPEG
    chb_ConvMark.Visible = False
    chb_EmbedFonts.Visible = False
    chb_Compress.Visible = False
    chb_Progressive.Visible = True
    chb_ExpMapGeoInfo.Visible = False
    cmb_PictSymb.Visible = False
    cmb_ImageComp.Visible = False
    cmb_ColorMode.Visible = True
    cmb_LayersAttrib.Visible = False
    chb_Progressive.Caption = "Progressive"
    Label4.Visible = False
    Label3.Visible = False
    Label5.Visible = True
    Label5.Caption = "Color mode:"
    cmb_ColorMode.Clear
    cmb_ColorMode.AddItem "24-bit True Color"
    cmb_ColorMode.AddItem "8-bit grayscale"
    cmb_ColorMode.ListIndex = 0
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = True
    Label9.Caption = "JPEG Quality"
    Label10.Visible = True
    Label17.Visible = False
    cmd_Transparency.Visible = False
    sld_ImageQ.Min = 0
    sld_ImageQ.Max = 100
    sld_ImageQ.SelStart = 100
    sld_ImageQ.Visible = True
    outputFormat = "jpg"
ElseIf cmb_ExpFormat.ListIndex = 2 Then     'PDF
    chb_ConvMark.Visible = True
    chb_EmbedFonts.Visible = True
    chb_Compress.Visible = True
    chb_Progressive.Visible = False
    chb_ExpMapGeoInfo.Visible = True
    chb_ConvMark.Value = 0
    chb_EmbedFonts.Value = 0
    chb_Compress.Value = 0
    chb_ExpMapGeoInfo.Value = 0
    cmb_PictSymb.Visible = True
    cmb_ImageComp.Visible = True
    cmb_ColorMode.Visible = True
    cmb_LayersAttrib.Visible = True
    chb_Compress.Caption = "Compress vector/text graphics"
    Label4.Visible = True
    Label3.Visible = True
    Label5.Visible = True
    Label5.Caption = "Destination Colorspace:"
    Label17.Visible = True
    cmb_ColorMode.Clear
    cmb_ColorMode.AddItem "RGB"
    cmb_ColorMode.AddItem "CMYK"
    cmb_ColorMode.Text = "RGB"
    colorMode = "RGB"
    cmb_LayersAttrib.Clear
    cmb_LayersAttrib.AddItem "None"
    cmb_LayersAttrib.AddItem "Export PDF Layers Only"
    cmb_LayersAttrib.AddItem "Export PDF Layers and Feature Attributes"
    cmb_LayersAttrib.Text = "Export PDF Layers Only"
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = True
    Label9.Caption = "Output Image Quality"
    Label10.Visible = True
    cmd_Transparency.Visible = False
    sld_ImageQ.Min = 1
    sld_ImageQ.Max = 5
    sld_ImageQ.SelStart = 3
    sld_ImageQ.Visible = True
    imageQuality = esriRasterOutputNormal
    cmb_ImageComp.Clear
    cmb_ImageComp.AddItem "None"
    cmb_ImageComp.AddItem "RLE"
    cmb_ImageComp.AddItem "LZW"
    cmb_ImageComp.AddItem "Deflate"
    cmb_ImageComp.Text = "Deflate"
    outputFormat = "pdf"
    chb_Compress.ToolTipText = "For PDF it compresses only vector and text portions of the map"
ElseIf cmb_ExpFormat.ListIndex = 3 Then     'PNG
    chb_ConvMark.Visible = False
    chb_EmbedFonts.Visible = False
    chb_Compress.Visible = False
    chb_Progressive.Visible = True
    chb_ExpMapGeoInfo.Visible = False
    cmb_PictSymb.Visible = False
    cmb_ImageComp.Visible = False
    cmb_ColorMode.Visible = True
    cmb_LayersAttrib.Visible = False
    chb_Progressive.Caption = "Interlaced"
    Label4.Visible = False
    Label3.Visible = False
    Label5.Visible = True
    Label5.Caption = "Color mode:"
    Label17.Visible = False
    cmb_ColorMode.Clear
    cmb_ColorMode.AddItem "24-bit True Color"
    cmb_ColorMode.AddItem "8-bit grayscale"
    cmb_ColorMode.ListIndex = 0
    Label7.Visible = True
    Label8.Visible = True
    Label9.Visible = False
    Label10.Visible = False
    sld_ImageQ.Visible = False
    cmd_Transparency.Visible = True
    outputFormat = "png"
ElseIf cmb_ExpFormat.ListIndex = 4 Then     'SVG
    chb_ConvMark.Visible = True
    chb_EmbedFonts.Visible = True
    chb_Compress.Visible = True
    chb_Progressive.Visible = False
    chb_ExpMapGeoInfo.Visible = False
    chb_ConvMark.Value = 0
    chb_EmbedFonts.Value = 0
    chb_Compress.Value = 0
    cmb_PictSymb.Visible = True
    cmb_ImageComp.Visible = False
    cmb_ColorMode.Visible = False
    cmb_LayersAttrib.Visible = False
    chb_Compress.Caption = "Compress document"
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = True
    Label9.Caption = "Output Image Quality"
    Label10.Visible = True
    Label17.Visible = False
    cmd_Transparency.Visible = False
    sld_ImageQ.Min = 1
    sld_ImageQ.Max = 5
    sld_ImageQ.SelStart = 3
    sld_ImageQ.Visible = True
    imageQuality = esriRasterOutputNormal
    Label4.Visible = False
    Label3.Visible = True
    Label5.Visible = False
    outputFormat = "svg"
    chb_Compress.ToolTipText = "For SVG changes the file extension to *.svgz"
ElseIf cmb_ExpFormat.ListIndex = 5 Then     'TIF
    chb_ConvMark.Visible = False
    chb_EmbedFonts.Visible = False
    chb_Compress.Visible = False
    chb_Progressive.Visible = False
    chb_ExpMapGeoInfo.Visible = False
    cmb_PictSymb.Visible = False
    cmb_ImageComp.Visible = True
    cmb_ColorMode.Visible = True
    cmb_LayersAttrib.Visible = False
    Label4.Visible = True
    Label3.Visible = False
    Label5.Visible = True
    Label17.Visible = False
    Label5.Caption = "Color mode:"
    cmb_ColorMode.Clear
    cmb_ColorMode.AddItem "24-bit True Color"
    cmb_ColorMode.AddItem "8-bit grayscale"
    cmb_ColorMode.ListIndex = 0
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = True
    Label9.Caption = "Deflate Quality"
    Label10.Visible = True
    cmd_Transparency.Visible = False
    sld_ImageQ.Min = 0
    sld_ImageQ.Max = 100
    sld_ImageQ.SelStart = 100
    sld_ImageQ.Visible = True
    sld_ImageQ.Enabled = True
    cmb_ImageComp.Clear
    cmb_ImageComp.AddItem "None"
    cmb_ImageComp.AddItem "LZW"
    cmb_ImageComp.AddItem "Deflate"
    cmb_ImageComp.AddItem "Pack Bits"
    cmb_ImageComp.AddItem "JPEG"
    cmb_ImageComp.Text = "Deflate"
    outputFormat = "tif"
End If

cmd_SetAsDefault.Enabled = True

End Sub


Private Sub cmb_ImageComp_Click()

If cmb_ImageComp.ListIndex = 0 Then
    If outputFormat = "tif" Then
        sld_ImageQ.Enabled = False
    End If
ElseIf cmb_ImageComp.ListIndex = 1 Then
    If outputFormat = "tif" Then
        sld_ImageQ.Enabled = False
    End If
ElseIf cmb_ImageComp.ListIndex = 2 Then
    If outputFormat = "tif" Then
        sld_ImageQ.Enabled = True
        Label9.Caption = "Deflate Quality"
    End If
ElseIf cmb_ImageComp.ListIndex = 3 Then
    If outputFormat = "tif" Then
        sld_ImageQ.Enabled = False
    End If
ElseIf cmb_ImageComp.ListIndex = 4 Then
    If outputFormat = "tif" Then
        sld_ImageQ.Enabled = True
        Label9.Caption = "JPEG Quality"
    End If
End If

cmd_SetAsDefault.Enabled = True

End Sub

Private Sub cmb_ImageComp_LostFocus()

If cmb_ImageComp.ListIndex = 0 Then
    If outputFormat = "tif" Then
        compType = esriTIFFCompressionNone
    Else
        compType = esriExportImageCompressionNone
    End If
ElseIf cmb_ImageComp.ListIndex = 1 Then
    If outputFormat = "tif" Then
        compType = esriTIFFCompressionLZW
    Else
        compType = esriExportImageCompressionRLE
    End If
ElseIf cmb_ImageComp.ListIndex = 2 Then
    If outputFormat = "tif" Then
        compType = esriTIFFCompressionDeflate
    Else
        compType = esriExportImageCompressionLZW
    End If
ElseIf cmb_ImageComp.ListIndex = 3 Then
    If outputFormat = "tif" Then
        compType = esriTIFFCompressionPackBits
    Else
        compType = esriExportImageCompressionDeflate
    End If
ElseIf cmb_ImageComp.ListIndex = 4 Then
    If outputFormat = "tif" Then
        compType = esriTIFFCompressionJPEG
    End If
End If


End Sub

Private Sub cmb_PictSymb_LostFocus()

If cmb_PictSymb.ListIndex = 0 Then
    pictSymb = esriPSORasterize
ElseIf cmb_PictSymb.ListIndex = 1 Then
    pictSymb = esriPSORasterizeIfRasterData
ElseIf cmb_PictSymb.ListIndex = 2 Then
    pictSymb = esriPSOVectorize
End If

cmd_SetAsDefault.Enabled = True

End Sub


Private Sub cmd_Execute_Click()

On Error GoTo err

Dim i As Integer, iRun As Integer
Dim mapDoc As IMapDocument
Dim pPageLayout As IPageLayout
Dim pActiveView As IActiveView
Dim arrFile()
Dim z As Integer, a As Integer
Dim totArcGisFile As Integer



'checking
If txb_ProgDir.Text = "" Then
    MsgBox "ATTENTION: please, select a directory!", vbExclamation, "Warning"
    Exit Sub
End If
If txb_dpi.Text = "" Then
    MsgBox "ATTENTION: please, set DPI resolution!", vbExclamation, "Warning"
    Exit Sub
End If
If IsNumeric(txb_dpi.Text) = False Then
    MsgBox "ATTENTION: DPI resolution is not correct!", vbExclamation, "Warning"
    Exit Sub
End If
If txb_dpi.Text < 96 Or txb_dpi.Text > 1200 Then
    MsgBox "ATTENTION: DPI resolution must be included between 96 and 1200 dpi!", vbExclamation, "Warning"
    Exit Sub
End If
If cmb_ExpFormat.ListIndex = 5 Then     'TIF
    If cmb_ImageComp.ListIndex = 2 And cmb_ColorMode.ListIndex = 1 Then
        MsgBox "ATTENTION: Deflate compression is not supported for 8-bit grayscale TIF!", vbExclamation, "Warning"
        Exit Sub
    End If
    If cmb_ImageComp.ListIndex = 4 And cmb_ColorMode.ListIndex = 1 Then
        MsgBox "ATTENTION: LZW compression is not supported for 8-bit grayscale TIF!", vbExclamation, "Warning"
        Exit Sub
    End If
End If
If opt_ExpSel.Value = True Then
    If UBound(agpSelFileArray) = 0 Then
        MsgBox "ATTENTION: no ArcGis projects selected!", vbExclamation, "Warning"
        Exit Sub
    End If
End If
If txb_Time2Run.Text <> "" And chb_Scheduled.Value = 0 Then
    MsgBox "ATTENTION: if you want to schedule the task you've to check the 'Scheduled' option! Otherwise delete the time in the textbox.", vbExclamation, "Warning"
    Exit Sub
End If
If txb_Custom.Text = "" And opt_schCustom.Value = True Then
    chb_Scheduled.Value = 0
    MsgBox "ATTENTION: you have to specify the interval (days) for scheduled task!", vbExclamation, "Warning"
    Exit Sub
End If



Dim FSO As FileSystemObject
Dim fls As Files
Dim flds As Folders
Dim f
Set FSO = CreateObject("Scripting.FileSystemObject")
Set fls = FSO.GetFolder(txb_ProgDir.Text).Files
Set flds = FSO.GetFolder(txb_ProgDir.Text).SubFolders

If chb_IncludeSub.Value = 1 Then
    If flds.Count = 0 Then
        MsgBox "ATTENTION: no subfolders found in the main folder!", vbExclamation, "Warning"
        Set FSO = Nothing
        Exit Sub
    End If
End If


If opt_ExpAll.Value = True Then
    'creates the array of ArcGis projects
    'in the main folder (and eventually in subfolders)
    '-------------------------
    Call LoopFolder.LoopFolderList
    '-------------------------
End If


If totAGP = 0 Then
    MsgBox "ATTENTION: no ArcGis projects found in the selected folder!", vbExclamation, "Warning"
    Exit Sub
End If


Dim answ As Integer
answ = MsgBox("Do you want to write a summary for the current export process?", vbYesNo, "Report")

cmd_Execute.Enabled = False
frame_Options.Enabled = False
Frame1.Enabled = False
Frame2.Enabled = False
Frame3.Enabled = False
Frame4.Enabled = False


iRun = 1


'scheduling task
'==============================================
If chb_Scheduled.Value = 1 Then
    frame_Schedule.Enabled = False
    progDate = DateValue(DTPicker1.Value) & " " & TimeValue(txb_Time2Run.Text)
    If opt_schCustom.Value = True Then
        customDays = txb_Custom.Text
    End If
    MsgBox "Batch Export Tool has been scheduled to run at: <<  " & progDate & "  >>", vbInformation, "Schedule task"
CheckTime2Run:
    If running = False Then
        If Dir(App.Path & "\Sch_Settings.ini") <> "" Then
            Kill App.Path & "\Sch_Settings.ini"
        End If
        Exit Sub
        Unload Me
    End If
    lbl_status.Caption = "- STANDBY -  Next run: " & progDate
    DoEvents
    If Now > progDate Then
        GoTo RunApp
    Else
        GoTo CheckTime2Run
    End If
End If
'==============================================



RunApp:

numOutput = 0
ReDim outputMaps(numOutput)

'set array for selected projects
If opt_ExpAll.Value = True Then
    totArcGisFile = UBound(agpAllFileArray) - 1
Else
    totArcGisFile = UBound(agpSelFileArray) - 1
End If


z = 1
For i = 0 To totArcGisFile

    Set mapDoc = New MapDocument
    
    If opt_ExpAll.Value = True Then
        lbl_status.Caption = "Exporting maps (" & z & " of " & totAGP & "). Please, wait..."
    Else
        lbl_status.Caption = "Exporting maps (" & z & " of " & UBound(agpSelFileArray) & "). Please, wait..."
    End If
    DoEvents
    
    If opt_ExpAll.Value = True Then
        a = InStrRev(agpAllFileArray(i), "\")
        fileAGP = Mid(agpAllFileArray(i), a + 1, Len(agpAllFileArray(i)) - a - 4)
        dirAGP = Mid(agpAllFileArray(i), 1, a - 1)
        mapDoc.Open agpAllFileArray(i)
    Else
        a = InStrRev(agpSelFileArray(i), "\")
        fileAGP = Mid(agpSelFileArray(i), a + 1, Len(agpSelFileArray(i)) - a - 4)
        mapDoc.Open dirAGP & "\" & fileAGP & "." & extAGP
    End If
    
    Set pPageLayout = mapDoc.PageLayout
    Set pActiveView = pPageLayout
    pActiveView.Activate GetDesktopWindow()

    'Set spatial reference for each dataframe of the iProject
    Dim j As Integer
    For j = 0 To mapDoc.MapCount - 1
        Dim pSpatialRef As ISpatialReference
        Set pSpatialRef = mapDoc.Map(j).SpatialReference
    Next j

    '----------------------------------------------------
    ExportActiveView2 mapDoc, pPageLayout, dirAGP, iRun
    '----------------------------------------------------

    mapDoc.Close
    Set mapDoc = Nothing
    z = z + 1
    
    If exportError = True Then
        Exit For
    End If
    
Next i


If answ = vbYes Then
    '------------------------------
    Call WriteReport(totArcGisFile + 1)
    '------------------------------
End If



If chb_Scheduled.Value = 1 Then
    If opt_schOnce.Value = True Then
        txb_Time2Run.Text = txb_Time2Run.Text
    ElseIf opt_schDaily.Value = True Then
        'Advance the date by one day
        DTPicker1.Value = Format$(Date + 1, "dd/mm/yyyy")
        progDate = DateValue(DTPicker1.Value) & " " & TimeValue(txb_Time2Run.Text)
    ElseIf opt_schCustom.Value = True Then
        'Advance the date by custom value
        DTPicker1.Value = Format$(Date + customDays, "dd/mm/yyyy")
        progDate = DateValue(DTPicker1.Value) & " " & TimeValue(txb_Time2Run.Text)
    End If
    'run again the batch export.....
    If opt_schDaily.Value = True Or opt_schCustom.Value = True Then
        iRun = iRun + 1
        GoTo CheckTime2Run
    End If
    'delete scheduling settings file at the end of the process
    If Dir(App.Path & "\Sch_Settings.ini") <> "" Then
        Kill App.Path & "\Sch_Settings.ini"
    End If
End If


frame_Options.Enabled = True
frame_Schedule.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Frame4.Enabled = True
chb_Scheduled.Value = 0
Set FSO = Nothing
cmd_Execute.Enabled = True
lbl_status.Caption = ""
DoEvents


If exportError = False Then
    If answ = vbYes Then
        If opt_schDaily.Value = True Or opt_schCustom.Value = True Then
            MsgBox "The report file has been saved in " & UCase(dirAGP) & ".", vbInformation, "EXPORT REPORT"
        ElseIf opt_schOnce.Value = True Then
            MsgBox "The report file 'exp_report_" & suffixReport & ".txt' has been saved in " & UCase(dirAGP) & ".", vbInformation, "EXPORT REPORT"
        End If
    End If
    MsgBox "Export finished at " & Format(Time, "Long Time") & "!", vbInformation, "Export"
Else
    MsgBox "PROGRAM STOPS!", vbCritical, "Error"
    exportError = False
End If



Exit Sub

err:
    MsgBox "ATTENTION: an error occurs in the routine 'Execute'!" & vbNewLine & _
            err.Description, vbCritical, "ERROR"
    exportError = True
    Exit Sub


End Sub


Public Sub ExportActiveView2(mapDoc As IMapDocument, pPageLayout As IPageLayout, folderAGP As String, iAGP As Integer)

On Error GoTo err

Dim hKey As Long

Dim pActiveView As IActiveView
Dim pEnv As IEnvelope
Dim exportFrame As tagRECT
Dim xMin As Double
Dim yMin As Double
Dim xMax As Double
Dim yMax As Double
Dim hdc As Long
Dim extFile As String
Dim pExportEPS As IExportPS
Dim pExportPDF As IExportPDF
Dim pExportPDF2 As IExportPDF2
Dim pExportTIF As IExportTIFF
Dim pExportJPG As IExportJPEG
Dim pExportSVG As IExportSVG
Dim pExportPNG As IExportPNG
Dim pExport As IExport
Dim exportVectorOpt As IExportVectorOptions
Dim exportVectorOptEx As IExportVectorOptionsEx
Dim pExportImage As IExportImage
Dim outImageName As String


Set pActiveView = mapDoc.ActiveView
Set pActiveView = mapDoc.PageLayout


'set the Best Output Image Quality (resample ratio = 1) for rasters (the same from menu File>Print)
Dim pGraphicsContainer As IGraphicsContainer
Dim pElement As IElement
Dim pOutputRasterSettings As IOutputRasterSettings
Dim pMapFrame As IMapFrame
Dim pTmpActiveView As IActiveView
If TypeOf pActiveView Is IMap Then
    Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
    pOutputRasterSettings.ResampleRatio = 1
ElseIf TypeOf pActiveView Is IPageLayout Then
    'assign ResampleRatio for PageLayout
    Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
    pOutputRasterSettings.ResampleRatio = 1
    'and assign ResampleRatio to the Maps in the PageLayout
    Set pGraphicsContainer = pActiveView
    pGraphicsContainer.Reset
    Set pElement = pGraphicsContainer.Next
    Do While Not pElement Is Nothing
        If TypeOf pElement Is IMapFrame Then
            Set pMapFrame = pElement
            Set pTmpActiveView = pMapFrame.Map
            Set pOutputRasterSettings = pTmpActiveView.ScreenDisplay.DisplayTransformation
            pOutputRasterSettings.ResampleRatio = 1
        End If
        DoEvents
        Set pElement = pGraphicsContainer.Next
    Loop
    Set pMapFrame = Nothing
    Set pGraphicsContainer = Nothing
    Set pTmpActiveView = Nothing
End If
Set pOutputRasterSettings = Nothing


If outputFormat = "jpg" Then
    Set pExportJPG = New ExportJPEG
    Set pExport = pExportJPG 'QI
    Set pExportImage = pExport
    pExportImage.ImageType = colorMode
    pExportJPG.Quality = sld_ImageQ.Value
    extFile = ".jpg"
ElseIf outputFormat = "tif" Then
    Set pExportTIF = New ExportTIFF
    Set pExport = pExportTIF 'QI
    Set pExportImage = pExport
    pExportImage.ImageType = colorMode
    extFile = ".tif"
ElseIf outputFormat = "pdf" Then
    Set pExportPDF = New ExportPDF
    Set pCSS = New ExportPDF
    If colorMode = "RGB" Then
        pCSS.Colorspace = esriExportColorspaceRGB
    Else
        pCSS.Colorspace = esriExportColorspaceCMYK
    End If
    Set pExport = pCSS  'QI
    Set pExport = pExportPDF 'QI
    Set pExportPDF2 = New ExportPDF
    pExportPDF2.ExportPDFLayersAndFeatureAttributes = layersAttrib
    pExportPDF2.ExportMeasureInfo = expGeoInfo
    Set pExport = pExportPDF2
    '-------------------------
    Call SetOutputQuality(pActiveView, imageQuality)
    '-------------------------
    extFile = ".pdf"
ElseIf outputFormat = "eps" Then
    Set pExportEPS = New ExportPS
    Set pCSS = New ExportPS
    If colorMode = "RGB" Then
        pCSS.Colorspace = esriExportColorspaceRGB
    Else
        pCSS.Colorspace = esriExportColorspaceCMYK
    End If
    Set pExport = pCSS  'QI
    Set pExport = pExportEPS 'QI
    '-------------------------
    Call SetOutputQuality(pActiveView, imageQuality)
    '-------------------------
    extFile = ".eps"
ElseIf outputFormat = "svg" Then
    Set pExportSVG = New ExportSVG
    Set pExport = pExportSVG 'QI
    '-------------------------
    Call SetOutputQuality(pActiveView, imageQuality)
    '-------------------------
    If chb_Compress.Value = 1 Then
        extFile = ".svgz"
    Else
        extFile = ".svg"
    End If
ElseIf outputFormat = "png" Then
    Set pExportPNG = New ExportPNG
    Set pExport = pExportPNG 'QI
    Set pExportImage = pExport
    pExportImage.ImageType = colorMode
    extFile = ".png"
End If



Set pEnv = New Envelope
dpi = txb_dpi.Text


'Setup the exporter
exportFrame = pActiveView.exportFrame
If pPageLayout.Page.Units = esriInches Then
    pEnv.PutCoords exportFrame.Left, exportFrame.Top, dpi * PageExtent(pPageLayout).UpperRight.X, _
                                dpi * PageExtent(pPageLayout).UpperRight.Y
ElseIf pPageLayout.Page.Units = esriCentimeters Then
    pEnv.PutCoords exportFrame.Left, exportFrame.Top, dpi * PageExtent(pPageLayout).UpperRight.X / 2.54, _
                                dpi * PageExtent(pPageLayout).UpperRight.Y / 2.54
End If


If opt_schOnce.Value = True Then
    outImageName = fileAGP & extFile
Else
    outImageName = fileAGP & "_" & iAGP & extFile
End If

With pExport
    .PixelBounds = pEnv
    If chb_OutFolder.Value = 0 Then
        .ExportFileName = folderAGP & "\" & outImageName
    Else
        .ExportFileName = txb_OutDir.Text & "\" & outImageName
    End If
    .Resolution = dpi
End With

'Recalc the export frame to handle the increased number of pixels
If outputFormat = "pdf" Then
    Set pEnv = pExport.PixelBounds
    pExportPDF.embedFonts = embedFonts
    pExportPDF.Compressed = mapCompress
    pExportPDF.ImageCompression = compType
    Set exportVectorOpt = pExport
    exportVectorOpt.PolygonizeMarkers = convMarkPoly
    Set exportVectorOptEx = pExport
    exportVectorOptEx.ExportPictureSymbolOptions = pictSymb
ElseIf outputFormat = "eps" Then
    Set pEnv = pExport.PixelBounds
    pExportEPS.embedFonts = embedFonts
    pExportEPS.ImageCompression = compType
    Set exportVectorOpt = pExport
    exportVectorOpt.PolygonizeMarkers = convMarkPoly
    Set exportVectorOptEx = pExport
    exportVectorOptEx.ExportPictureSymbolOptions = pictSymb
ElseIf outputFormat = "jpg" Then
    Set pEnv = pExport.PixelBounds
    pExportJPG.ProgressiveMode = progMode
ElseIf outputFormat = "tif" Then
    Set pEnv = pExport.PixelBounds
    pExportTIF.CompressionType = compType
ElseIf outputFormat = "svg" Then
    Set pEnv = pExport.PixelBounds
    pExportSVG.embedFonts = embedFonts
    pExportSVG.Compressed = mapCompress
    Set exportVectorOpt = pExport
    exportVectorOpt.PolygonizeMarkers = convMarkPoly
    Set exportVectorOptEx = pExport
    exportVectorOptEx.ExportPictureSymbolOptions = pictSymb
ElseIf outputFormat = "png" Then
    Set pEnv = pExport.PixelBounds
    pExportPNG.InterlaceMode = progMode
    If traspCol = True Then
        pExportPNG.TransparentColor = pOutColor
    End If
End If

pEnv.QueryCoords xMin, yMin, xMax, yMax
exportFrame.Left = xMin
exportFrame.Top = yMin
exportFrame.Right = xMax
exportFrame.Bottom = yMax


'///////////////////////////////////////////////////
'Do the export
hdc = pExport.StartExporting
    pActiveView.Output hdc, dpi, exportFrame, Nothing, Nothing
pExport.FinishExporting
'///////////////////////////////////////////////////


Dim a As Integer
a = InStrRev(folderAGP, "\")
Dim originalFolder As String
originalFolder = Mid(folderAGP, a + 1)

'create an array with output filenames
ReDim Preserve outputMaps(numOutput + 1)
outputMaps(numOutput) = "(" & originalFolder & ") >> " & outImageName
numOutput = numOutput + 1


Exit Sub

err:
    MsgBox "ATTENTION: an error occurs in the routine 'ExportActiveView2'!" & vbNewLine & _
            err.Description, vbCritical, "ERROR"
    exportError = True
    Exit Sub

End Sub



Private Sub SetOutputQuality(ByVal docActiveView As IActiveView, ByVal iResampleRatio As Long)

Dim oiqMap As IMap
Dim docGraphicsContainer As IGraphicsContainer
Dim docElement As IElement
Dim docOutputRasterSettings As IOutputRasterSettings
Dim oiqMapFrame As IMapFrame
Dim TmpActiveView As IActiveView

If TypeOf docActiveView Is IMap Then

    Set docOutputRasterSettings = docActiveView.ScreenDisplay.DisplayTransformation
    docOutputRasterSettings.ResampleRatio = iResampleRatio
    
ElseIf TypeOf docActiveView Is IPageLayout Then

    'assign ResampleRatio for PageLayout
    Set docOutputRasterSettings = docActiveView.ScreenDisplay.DisplayTransformation
    docOutputRasterSettings.ResampleRatio = iResampleRatio

    'and assign ResampleRatio to the Maps in the PageLayout
    Set docGraphicsContainer = docActiveView
    docGraphicsContainer.Reset
    Set docElement = docGraphicsContainer.Next
    Do While Not docElement Is Nothing
        If TypeOf docElement Is IMapFrame Then
            Set oiqMapFrame = docElement
            Set TmpActiveView = oiqMapFrame.Map
            Set docOutputRasterSettings = TmpActiveView.ScreenDisplay.DisplayTransformation
            docOutputRasterSettings.ResampleRatio = iResampleRatio
        End If
        Set docElement = docGraphicsContainer.Next
    Loop
    
    Set oiqMap = Nothing
    Set oiqMapFrame = Nothing
    Set docGraphicsContainer = Nothing
    Set TmpActiveView = Nothing
    
End If

Set docOutputRasterSettings = Nothing

End Sub


Private Sub WriteReport(totMaps As Integer)

On Error GoTo err

Dim FSO As New FileSystemObject
Dim repFile
Dim ii As Integer, i As Integer


If chb_Scheduled.Value = 1 Then
    suffixReport = Replace(progDate, " ", "_")
    suffixReport = Replace(suffixReport, "/", ".")
Else
    suffixReport = Format(Date, "dd.mm.yyyy") & "_" & Time
End If


Set FSO = CreateObject("Scripting.FileSystemObject")


If chb_IncludeSub.Value = 1 Then
    dirAGP = txb_ProgDir.Text
End If
If chb_OutFolder.Value = 1 Then
    dirAGP = txb_OutDir.Text
End If



Set repFile = FSO.OpenTextFile(dirAGP & "\exp_report_" & suffixReport & ".txt", ForWriting, True)
    repFile.WriteLine "BATCH EXPORT PROCESS - Report"
    repFile.WriteLine "-----------------------------"
    repFile.WriteLine
    repFile.WriteLine "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    If chb_OutFolder.Value = 1 Then
        repFile.WriteLine "Destination folder: " & UCase(dirAGP)
    Else
        If chb_IncludeSub.Value = 1 Then
            repFile.WriteLine "Main destination folder: " & UCase(dirAGP)
        Else
            repFile.WriteLine "Destination folder: " & UCase(dirAGP)
        End If
    End If
    repFile.WriteLine "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    repFile.WriteLine "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    If opt_ExpAll.Value = True Then
        If chb_IncludeSub.Value = 1 Then
            repFile.WriteLine "Total Exported Maps: " & totMaps & " (ALL - Subfolders included)"
        Else
            repFile.WriteLine "Total Exported Maps: " & totMaps & " (ALL)"
        End If
    Else
        repFile.WriteLine "Total Exported Maps: " & totMaps & " (SUBSET)"
    End If
    repFile.WriteLine "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    repFile.WriteLine
    repFile.WriteLine "ArcGis project extension:             " & UCase(extAGP)
    repFile.WriteLine "Output format:                        " & UCase(outputFormat)
    repFile.WriteLine "Output resolution:                    " & dpi & " dpi"
    repFile.WriteLine
    
    If outputFormat = "eps" Or outputFormat = "pdf" Or outputFormat = "tif" Then
        repFile.WriteLine "Image compression:                    " & cmb_ImageComp.Text
    End If
    
    If outputFormat = "eps" Or outputFormat = "pdf" Or outputFormat = "svg" Then
        repFile.WriteLine "Picture symbol:                       " & cmb_PictSymb.Text
    End If
    
    If outputFormat = "jpg" Or outputFormat = "png" Or outputFormat = "tif" Then
        repFile.WriteLine "Color mode:                           " & cmb_ColorMode.Text
    ElseIf outputFormat = "eps" Or outputFormat = "pdf" Then
        repFile.WriteLine "Destination colorspace:               " & cmb_ColorMode.Text
    End If
    
    If outputFormat = "pdf" Then
        repFile.WriteLine "Layer and Attributes:                 " & cmb_LayersAttrib.Text
    End If
    
    If outputFormat = "png" Then
        repFile.WriteLine "Transparency color:                   " & traspCol
    End If
    
    If outputFormat = "eps" Or outputFormat = "pdf" Or outputFormat = "svg" Then
        repFile.WriteLine "Embed all documents fonts:            " & embedFonts
    End If
    
    If outputFormat = "eps" Or outputFormat = "pdf" Or outputFormat = "svg" Then
        repFile.WriteLine "Convert marker symbols to polygons:   " & convMarkPoly
    End If
    
    If outputFormat = "pdf" Then
        repFile.WriteLine "Compress vector/text graphics:        " & mapCompress
    ElseIf outputFormat = "svg" Then
        repFile.WriteLine "Compress document:                    " & mapCompress
    End If
    
    If outputFormat = "jpg" Then
        repFile.WriteLine "Progressive:                          " & progMode
    ElseIf outputFormat = "png" Then
        repFile.WriteLine "Interlaced:                           " & progMode
    End If
    
    If outputFormat = "eps" Or outputFormat = "pdf" Or outputFormat = "svg" Then
        If sld_ImageQ.Value = 1 Or sld_ImageQ.Value = 2 Then
            repFile.WriteLine "Output Image Quality:                 Draft"
        ElseIf sld_ImageQ.Value = 3 Then
            repFile.WriteLine "Output Image Quality:                 Normal"
        ElseIf sld_ImageQ.Value = 4 Or sld_ImageQ.Value = 5 Then
            repFile.WriteLine "Output Image Quality:                 Best"
        End If
    ElseIf outputFormat = "jpg" Then
        repFile.WriteLine "JPEG Quality:                         " & sld_ImageQ.Value
    ElseIf outputFormat = "tif" Then
        If cmb_ImageComp.Text = "Deflate" Then
            repFile.WriteLine "Deflate Quality:                      " & sld_ImageQ.Value
        ElseIf cmb_ImageComp.Text = "JPEG" Then
            repFile.WriteLine "JPEG Quality:                         " & sld_ImageQ.Value
        End If
    End If
    
    If outputFormat = "pdf" Then
        repFile.WriteLine "Export Map Georeference Information:  " & expGeoInfo
    End If
    
    repFile.WriteLine
    repFile.WriteLine "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
    repFile.WriteLine
    
    repFile.WriteLine
    repFile.WriteLine "-- Exported Maps: --"
    repFile.WriteLine
    
    i = 0
    For i = 0 To UBound(outputMaps) - 1
        repFile.WriteLine "   " & UCase(outputMaps(i)) & vbCr
    Next
    
    repFile.Close
Set repFile = Nothing


Set FSO = Nothing


Exit Sub

err:
    MsgBox "ATTENTION: an error occurs in the routine 'WriteReport'!" & vbNewLine & _
            err.Description, vbCritical, "ERROR"
    exportError = True
    Exit Sub

End Sub


Function PageExtent(pPageLayout As IPageLayout) As IEnvelope

Dim dWidth As Double, dHeight As Double

pPageLayout.Page.QuerySize dWidth, dHeight

Dim pEnv As IEnvelope
Set pEnv = New Envelope
pEnv.PutCoords 0#, 0#, dWidth, dHeight

Set PageExtent = pEnv

End Function


Private Sub cmd_SelProjects_Click()

'creates the array of selected ArcGis projects in the selected folder

On Error GoTo err

If txb_ProgDir.Text = "" Then
    MsgBox "ATTENTION: please, select a folder!", vbExclamation, "Warning"
    Exit Sub
End If

Dim pGxDialog As IGxDialog
Dim pFileFilter As IGxFileFilter
Dim pGxCat As IGxCatalog
Dim pHandle As OLE_HANDLE
Dim pSelection As IGxSelection  '

lbl_status.Caption = "Opening..."
DoEvents

Set pGxDialog = New GxDialog
Set pGxCat = pGxDialog.InternalCatalog

Set pSelection = pGxCat.Selection   '
Set pFileFilter = pGxCat.FileFilter


Dim fileExt As String
If extAGP = "mxd" Then
    fileExt = "mxt"
Else
    fileExt = "mxd"
End If

'load projects
If (pGxCat.FileFilter.FindFileType(fileExt) = -1) Then
    pFileFilter.AddFileType fileExt, "File type: " & UCase(fileExt) & "", ""
End If

'filter projects
Dim pObjFilter As IGxObjectFilter
Set pObjFilter = New GxFilterMaps

Set pGxDialog.ObjectFilter = pObjFilter

pGxDialog.AllowMultiSelect = True
pGxDialog.Title = "Select ArcGis projects (" & UCase(extAGP) & ") to export..."
pGxDialog.StartingLocation = txb_ProgDir.Text

lbl_status.Caption = ""
DoEvents

Dim pGxEnum As IEnumGxObject
pGxDialog.DoModalOpen pHandle, pGxEnum


Dim pGxObject As IGxObject

totAGPFiles = 0
ReDim agpSelFileArray(totAGPFiles)

Set pGxEnum = pSelection.SelectedObjects
pGxEnum.Reset
Set pGxObject = pGxEnum.Next

Do While Not pGxObject Is Nothing
    ReDim Preserve agpSelFileArray(totAGPFiles + 1)
    agpSelFileArray(totAGPFiles) = pGxObject.FullName
    dirAGP = Mid(pGxObject.FullName, 1, Len(pGxObject.FullName) - Len(pGxObject.Name) - 1)
    totAGPFiles = totAGPFiles + 1
    Set pGxObject = pGxEnum.Next
Loop

pGxDialog.InternalCatalog.Close


txb_ProgDir.Text = dirAGP


Frame4.Caption = "Projects to export: " & totAGPFiles



Exit Sub

err:
    MsgBox "ATTENTION: an error occurs on click event of SelProjects command button!" & vbNewLine & _
            err.Description, vbCritical, "ERROR"
    exportError = True
    Exit Sub


End Sub

Private Sub cmd_Transparency_Click()


'Set the initial color to be diaplyed when the dialog opens
Dim pOutColor As IColor
Dim pColor As IColor
Set pColor = New RgbColor
pColor.RGB = 255 'Red

Dim bColorSet As Boolean
Dim pSelector As IColorSelector
Set pSelector = New ColorSelector
pSelector.Color = pColor

bColorSet = pSelector.DoModal(0)

' Display the dialog
If bColorSet Then
    Set pOutColor = pSelector.Color
    Label8.BackColor = pOutColor.RGB
    Label8.Caption = ""
    traspCol = True
Else
    Label8.Caption = "No color"
    traspCol = False
End If

cmd_SetAsDefault.Enabled = True

End Sub

Private Sub Form_Load()

    
'Perform license initialization on a system

'It will check the required licenses and keep them checked out
Dim licenseStatus As esriLicenseStatus
licenseStatus = CheckOutLicenses(esriLicenseProductCodeArcView)

'Take a look at the licenseStatus to see if it failed
'Not licensed
If (licenseStatus = esriLicenseNotLicensed) Then
  'MsgBox "You are not licensed to run this product.", vbCritical
  'Unload Me
  'Exit Sub
'The licenses needed are currently in use
ElseIf (licenseStatus = esriLicenseUnavailable) Then
  MsgBox "There are insuficient licenses to run.", vbCritical
  Unload Me
  Exit Sub
'The licenses unexpected license failure
ElseIf (licenseStatus = esriLicenseFailure) Then
  MsgBox "Unexpected license failure: please contact your administrator.", vbCritical
  Unload Me
  Exit Sub
'Already initialized (Initialization can only occur once)
ElseIf (licenseStatus = esriLicenseAlreadyInitialized) Then
  MsgBox "Your license has already been initialized, please check your implementation.", vbCritical
  Unload Me
  Exit Sub
End If


running = True


If closeApp = True Then
    If Dir(App.Path & "\Sch_Settings.ini") <> "" Then
        '-------------------------------
        Call Settings.LoadSavedSettings
        '-------------------------------
    End If
Else
    If Dir(App.Path & "\Settings.ini") <> "" Then
        '-------------------------------
        Call Settings.LoadSavedSettings
        '-------------------------------
    Else
        '-------------------------------
        Call Settings.DefaultSettings
        '-------------------------------
    End If
End If


'to minimize in the tray
'the form must be fully visible before calling Shell_NotifyIcon
Me.Show
Me.Refresh
With nid
    'The length of the NOTIFYICONDATA type
    .cbSize = Len(nid)
    'hWnd of the form
    .hwnd = Me.hwnd
    'uID is not used by VB, so it's set to a Null value
    .uId = vbNull
    'It will have message handling and a tooltip
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    'The icon we're placing will send messages to the MouseMove event
    .uCallBackMessage = WM_MOUSEMOVE
    'A reference to the form's icon
    .hIcon = Me.Icon
    'Tooltip string delimited with a null character
    .szTip = "Batch Export Tool 1.6" & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid


End Sub

Private Sub mPopExit_Click()

'called when user clicks the popup menu Exit command
Unload Me

End Sub
      
Private Sub mPopRestore_Click()

'called when the user clicks the popup menu Restore command
Me.WindowState = vbNormal
Result = SetForegroundWindow(Me.hwnd)
Me.Show

End Sub


Private Sub opt_ExpAll_Click()

If opt_ExpAll.Value = True Then
    opt_ExpAll.FontBold = True
    opt_ExpAll.ForeColor = RGB(218, 27, 8)
    cmd_SelProjects.Enabled = False
    If includeSub = True Then
        chb_IncludeSub.Value = 1
    End If
    If txb_ProgDir.Text <> "" Then
        '-------------------------
        Call LoopFolder.LoopFolderList
        '-------------------------
        If totAGP > 0 Then
            Frame4.Caption = "Projects to export: ALL (" & totAGP & ")"
        Else
            Frame4.Caption = "Projects to export: ALL (0)"
        End If
    Else
        Frame4.Caption = "Projects to export: ALL (0)"
    End If
End If

opt_ExpSel.FontBold = False
opt_ExpSel.ForeColor = RGB(0, 0, 0)

End Sub

Private Sub opt_ExpSel_Click()

If opt_ExpSel.Value = True Then
    opt_ExpSel.FontBold = True
    opt_ExpSel.ForeColor = RGB(218, 27, 8)
    cmd_SelProjects.Enabled = True
    If chb_IncludeSub.Value = 1 Then
        MsgBox "You can't include subfolders with 'Selected' projects to export option checked!", vbExclamation, "Warning"
        chb_IncludeSub.Value = 1
        opt_ExpAll.Value = True
        includeSub = False
        Exit Sub
    End If
    chb_IncludeSub.FontBold = False
    chb_IncludeSub.ForeColor = RGB(0, 0, 0)
    Frame4.Caption = "Projects to export: " & totAGPFiles
End If

opt_ExpAll.FontBold = False
opt_ExpAll.ForeColor = RGB(0, 0, 0)

End Sub

Private Sub optMxd_Click()

If optMxd.Value = True Then
    optMxd.FontBold = True
    optMxd.ForeColor = RGB(218, 27, 8)
    extAGP = "mxd"
    If txb_ProgDir.Text <> "" And opt_ExpAll.Value = True Then
        '-------------------------
        Call LoopFolder.LoopFolderList
        '-------------------------
        If totAGP > 0 Then
            MsgBox "You're going to export  << " & totAGP & " >>  ArcGis " & UCase(extAGP) & " projects.", vbInformation, "Projects to export"
            Frame4.Caption = "Projects to export: ALL (" & totAGP & ")"
        Else
            Frame4.Caption = "Projects to export: ALL (0)"
        End If
    End If
End If

optMxt.FontBold = False
optMxt.ForeColor = RGB(0, 0, 0)

End Sub

Private Sub optMxt_Click()

If optMxt.Value = True Then
    optMxt.FontBold = True
    optMxt.ForeColor = RGB(218, 27, 8)
    extAGP = "mxt"
    If txb_ProgDir.Text <> "" And opt_ExpAll.Value = True Then
        '-------------------------
        Call LoopFolder.LoopFolderList
        '-------------------------
        If totAGP > 0 Then
            MsgBox "You're going to export  << " & totAGP & " >>  ArcGis " & UCase(extAGP) & " projects.", vbInformation, "Projects to export"
            Frame4.Caption = "Projects to export: ALL (" & totAGP & ")"
        Else
            Frame4.Caption = "Projects to export: ALL (0)"
        End If
    End If
End If

optMxd.FontBold = False
optMxd.ForeColor = RGB(0, 0, 0)

End Sub

Private Sub sld_ImageQ_Click()

If outputFormat = "svg" Or outputFormat = "pdf" Or outputFormat = "eps" Then
    If sld_ImageQ.Value = 1 Then
        imageQuality = esriRasterOutputDraft
    ElseIf sld_ImageQ.Value = 2 Then
        imageQuality = esriRasterOutputDraft
    ElseIf sld_ImageQ.Value = 3 Then
        imageQuality = esriRasterOutputNormal
    ElseIf sld_ImageQ.Value = 4 Then
        imageQuality = esriRasterOutputBest
    ElseIf sld_ImageQ.Value = 5 Then
        imageQuality = esriRasterOutputBest
    End If
End If

cmd_SetAsDefault.Enabled = True

End Sub

Private Sub txb_Custom_LostFocus()

If txb_Custom.Text <> "" Then
    If txb_Custom.Text = 0 Or txb_Custom.Text = 1 Then
        MsgBox "ATTENTION: scheduled interval must be greater than 1!", vbExclamation, "Warning"
    End If
End If

End Sub

Private Sub txb_dpi_Change()

cmd_SetAsDefault.Enabled = True

End Sub

Private Sub txb_Time2Run_Change()

If txb_Time2Run.Text <> "" Then
    chb_Scheduled.Enabled = True
Else
    chb_Scheduled.Enabled = False
End If

End Sub
