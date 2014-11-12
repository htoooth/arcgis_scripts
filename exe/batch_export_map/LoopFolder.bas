Attribute VB_Name = "LoopFolder"
Option Explicit

Public extAGP As String
Public totAGP As Integer
Public agpSubFolders As Boolean
Public agpAllFileArray()
Dim arcGisFile As String


Public Sub LoopFolderList()

Dim FSO As Object, file As Object, folder As Object, subfolder As Object, s As String

Set FSO = CreateObject("Scripting.FileSystemObject")
Set folder = FSO.GetFolder(Export_Map.txb_ProgDir.Text)


Export_Map.cmd_Execute.Enabled = False
Export_Map.lbl_status.Caption = "Looking for ArcGis projects..."
DoEvents

agpSubFolders = False

totAGP = 0
ReDim agpAllFileArray(totAGP)

For Each file In folder.Files
    If Right(UCase(file.Name), 3) = UCase(extAGP) Then
        's = s & file.Name & " (" & folder.Name & ")" & vbCr
        agpAllFileArray(totAGP) = file
        ReDim Preserve agpAllFileArray(totAGP + 1)
        totAGP = totAGP + 1
    End If
Next file

If Export_Map.chb_IncludeSub.Value = 1 Then
    For Each subfolder In folder.SubFolders
        Call LoopSubFolderList(subfolder, s)
    Next subfolder
End If


Export_Map.cmd_Execute.Enabled = True
Export_Map.lbl_status.Caption = ""
DoEvents
    
    
Set file = Nothing
Set subfolder = Nothing
Set FSO = Nothing
    
End Sub



Private Sub LoopSubFolderList(fld As Object, ByRef str As String)

Dim fil As Object, subfld As Object, arr() As Variant
    
For Each fil In fld.Files
    agpSubFolders = True
    If Right(UCase(fil.Name), 3) = UCase(extAGP) Then
        'str = str & fil.Name & " (" & fld.Name & ")" & vbCr
        agpAllFileArray(totAGP) = fil
        ReDim Preserve agpAllFileArray(totAGP + 1)
        totAGP = totAGP + 1
    End If
Next fil
    
For Each subfld In fld.SubFolders
    Call LoopSubFolderList(subfld, str)
Next subfld
    
Set fil = Nothing
Set subfld = Nothing
    
End Sub



