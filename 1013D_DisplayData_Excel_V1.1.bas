Attribute VB_Name = "Modul1"
' FNIRSI-1013D WAV Viewer for Excel VBA
' Data Format: original Firmware .wav Format

Option Explicit

Public Const Version = "V1.10"

Dim Bytes     'The whole file in a Bytes Array
Dim VScaleArr
Dim TScaleArr
Dim ChRow0 As Long
Dim Ch1Col As Long
Dim Ch2Col As Long
Dim TimCol As Long

Dim FileName As String
Dim FileIsOpen As Boolean

Sub LoadData()
    VScaleArr = Array(5, 2.5, 1, 0.5, 0.2, 0.1, 0.05)
    TScaleArr = Array(50, 20, 10, 5, 2, 1, 0.5, 0.2, 0.1, 0.05, 0.02, 0.01, 0.005, 0.002, 0.001, 0.0005, 0.0002, 0.0001, 0.00005, 0.00002, 0.00001, 0.000005, 0.000002, 0.000001, 0.0000005, 0.0000002, 0.0000001, 0.00000005, 0.00000002, 0.00000001)
    On Error GoTo ErrHandler
    With CreateObject("ADODB.Stream")  ' load file
        .Open
        .Type = 1  ' adTypeBinary
        .LoadFromFile FileName
        Bytes = .Read
        .Close
    End With
    TimCol = Range("TimeVals").Column
    Ch1Col = Range("Ch1Vals").Column
    Ch2Col = Range("Ch2Vals").Column
    ChRow0 = Range("Ch1Vals").Row
    Range("Version") = Version
    Range(Cells(ChRow0 + 1, Ch1Col), Cells(ChRow0 + 1500, Ch2Col)).ClearContents
    Cells(ChRow0 + 1, Ch1Col) = " "
    Cells(ChRow0 + 1, Ch2Col) = " "
    If (Range("ChkChan2") = False) Then Range("ChkChan1") = True
    Application.Calculation = xlCalculationManual
    If (Range("ChkChan1") = True) Then PlotData 1, 1000, 3000  ' plot Channel
    If (Range("ChkChan2") = True) Then PlotData 2, 4000, 3000
    Application.Calculation = xlCalculationAutomatic
    FileIsOpen = True
Exit Sub
ErrHandler:
    Call MsgBox("Can't open file", vbCritical Or vbOKOnly, Error)
End Sub


Sub PlotData(ch, dataStart, dataSize)
    
    Dim index As Integer
    Dim vVal As Integer
    Dim vMult As Single
    Dim tMult As Single
    Dim probe As Single
    Dim rowcnt As Integer
    Dim dataEnd As Integer
    
    probe = 10 ^ Bytes(10 + (ch - 1) * 10) ' Probe X
    Range("ProbeCh1").Offset(ch - 1) = probe
    vMult = VScaleArr(Bytes(4 + (ch - 1) * 10)) * probe ' Vertical scale
    Range("VDivCh1").Offset(ch - 1) = vMult
    tMult = TScaleArr(Bytes(22))           ' Time scale
    Range("TDivCh1").Offset(ch - 1) = tMult
    
    rowcnt = ChRow0
    dataEnd = dataStart + dataSize - 2
    For index = dataStart To dataEnd Step 2
        vVal = (Bytes(index + 1) * 256 + Bytes(index)) - 200
        rowcnt = rowcnt + 1
        Cells(rowcnt, TimCol) = (rowcnt - 2) * tMult / 50    ' Horizontal
        Cells(rowcnt, Ch1Col + ch - 1) = vVal * vMult / 50 * probe ' Vertical
    Next
End Sub



Sub Export_CSV()
    On Error GoTo ErrHandler
    If Not FileIsOpen Then
        Call MsgBox("Please open file first", vbCritical Or vbOKOnly, Error)
        Exit Sub
    End If
    Dim rowcnt As Integer
    Dim f As Long
    f = FreeFile()
    FileName = Range("FileName")
    Open FileName & ".csv" For Output As #f
    For rowcnt = ChRow0 + 1 To ChRow0 + 1500
        Print #f, CStr(Cells(rowcnt, TimCol)) & ";" & CStr(Cells(rowcnt, Ch1Col)) & ";" & CStr(Cells(rowcnt, Ch2Col))
    Next
    Close #f
Exit Sub
ErrHandler:
    Call MsgBox("Error saving csv file", vbCritical Or vbOKOnly, Error)
End Sub



Sub Chan_Select()
    FileName = Range("FileName")
    If FileName <> "" Then LoadData
End Sub


Sub Open_File()
    FileName = Range("FileName")
    FileName = Get_FileName
    If FileName <> "" Then
        LoadData
        Range("FileName") = FileName
    End If
End Sub


Public Function Get_FileName() As String
    Dim f As Office.FileDialog
    Set f = Application.FileDialog(msoFileDialogFilePicker)
    
    With f
        .Title = "Open File"                      'Fenstertitel
        .AllowMultiSelect = False                 'Nur eine Datei auswählbar
        .ButtonName = "Open"                      'Button Beschriftung
        .Filters.Clear                            'erst alle Filter löschen
        .Filters.Add "FNIRSI 1013D", "*.wav"      'dann eigene anlegen
        '.FilterIndex = 1                          'einen Filter vorselektieren
        .InitialFileName = FileName               'Startverzeichnis
        .Show
    End With
    
    If f.SelectedItems.Count > 0 Then
        Get_FileName = f.SelectedItems(1)
    End If
End Function


