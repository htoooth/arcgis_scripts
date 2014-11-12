Attribute VB_Name = "RunLater"


Public Sub ScheduledRun()

Dim FSO As New FileSystemObject
Dim repFile
Dim ii As Integer, i As Integer
Dim suffixReport As String

suffixReport = Format(Date, "dd.mm.yyyy") & "_" & Time

Set FSO = CreateObject("Scripting.FileSystemObject")

Set repFile = FSO.OpenTextFile(dirAGP & "\scdjob_" & suffixReport & ".txt", ForWriting, True)

    repFile.WriteLine Export_Map.txb_ProgDir.Text
    repFile.WriteLine Export_Map.chb_IncludeSub.Value
    repFile.WriteLine Export_Map.chb_OutFolder.Value
    repFile.WriteLine Export_Map.txb_OutDir
    repFile.WriteLine Export_Map.optMxd.Value
    repFile.WriteLine Export_Map.optMxt.Value
    repFile.WriteLine Export_Map.opt_ExpAll.Value
    repFile.WriteLine Export_Map.opt_ExpSel.Value
    repFile.WriteLine Export_Map.txb_dpi.Text
    repFile.WriteLine Export_Map.cmb_ExpFormat.Text
    repFile.WriteLine Export_Map.cmb_ImageComp.Text
    repFile.WriteLine Export_Map.cmb_PictSymb.Text
    repFile.WriteLine Export_Map.cmb_ColorMode.Text
    repFile.WriteLine Export_Map.chb_EmbedFonts.Value
    repFile.WriteLine Export_Map.chb_ConvMark.Value
    repFile.WriteLine Export_Map.chb_Compress.Value
    repFile.WriteLine Export_Map.chb_Progressive.Value
    repFile.WriteLine Export_Map.sld_ImageQ.Value
    repFile.WriteLine
    repFile.WriteLine
    repFile.WriteLine



repFile.Close
Set repFile = Nothing
Set FSO = Nothing

End Sub
