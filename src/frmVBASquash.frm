VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVBASquash 
   Caption         =   "vbaSquash"
   ClientHeight    =   4668
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9168.001
   OleObjectBlob   =   "frmVBASquash.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVBASquash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Squash As vbaSquash

Private OldData() As Byte
Private NewData() As Byte
Private CompressTheFile As Boolean
Private CompressionAlgorithm As COMPRESS_ALGORITHM_ENUM
Private IsLoadingFile As Boolean

Private AlgoIds As Variant

Private Sub UserForm_Initialize()
  Set m_Squash = New vbaSquash

  AlgoIds = Array(MSZIP, XPRESS, XPRESS_HUFF, LZMS, RTL_LZNT1, RTL_XPRESS, RTL_XPRESS_HUFFMAN)
  cboxMethod.List = Split("MS-ZIP|XPRESS|XPRESS HUFF|LZMS|RTL LZNT1|RTL XPRESS|RTL XPRESS HUFF", "|")
  cboxMethod.ListIndex = 1

  btnCompress.Enabled = False
  btnSave.Enabled = False
  btnClear.Enabled = False

  txtCompSize.Locked = True
  txtPercentage.Locked = True
End Sub

Private Sub UserForm_Terminate()
  Set m_Squash = Nothing
End Sub

Private Sub txtFlename_Change()
  btnClear.Enabled = (Len(txtFlename.Text) > 0)
End Sub

Private Sub txtFlename_AfterUpdate()
  If IsLoadingFile Then Exit Sub
  IsLoadingFile = True
  LoadFile txtFlename.Text
  IsLoadingFile = False
End Sub

Private Sub btnBrowseSelect_Click()
  Dim fPath As Variant
  fPath = Application.GetOpenFilename("All Files,*.*", , "Select a file to process")
  If fPath <> False Then
    If IsLoadingFile Then Exit Sub
    IsLoadingFile = True
    txtFlename.Text = fPath
    LoadFile fPath
    IsLoadingFile = False
  End If
End Sub

Private Sub cboxMethod_Change()
  Erase NewData
  txtCompSize.Text = vbNullString
  txtPercentage.Text = vbNullString
  btnSave.Enabled = False
End Sub

Private Sub btnCompress_Click()
  txtCompSize.Text = vbNullString
  txtPercentage.Text = vbNullString
  btnSave.Enabled = False
  Erase NewData

  If cboxMethod.ListIndex < 0 Then Exit Sub
  CompressionAlgorithm = AlgoIds(cboxMethod.ListIndex)

  If CompressTheFile Then
    NewData = m_Squash.CompressBytes(OldData, CompressionAlgorithm)
  Else
    ' The header is authoritative: let the class detect the algorithm (and, for RTL
    ' data, the original size) rather than trusting the combo selection.
    NewData = m_Squash.DecompressBytes(OldData)
  End If

  ' A failed (de)compression returns an empty array; UBound on it would raise error 9.
  If m_Squash.CheckArray(NewData) = 0 Then
    If CompressTheFile Then txtCompSize.Text = "Failed" Else txtDecompSize.Text = "Failed"
    Exit Sub
  End If

  Dim OriginalSize As Long: OriginalSize = m_Squash.CheckArray(OldData)
  Dim newSize As Long: newSize = m_Squash.CheckArray(NewData)

  txtDecompSize.Text = IIf(CompressTheFile, OriginalSize, newSize)
  txtCompSize.Text = IIf(CompressTheFile, newSize, OriginalSize)
  txtPercentage.Text = Format(100 * (newSize / OriginalSize), "0.0") & "%"

  btnSave.Enabled = True
End Sub

Private Sub btnSave_Click()
  If Len(txtSaveAs.Text) = 0 Then Exit Sub
  If m_Squash.CheckArray(NewData) = 0 Then Exit Sub

  If m_Squash.WriteFile(txtSaveAs.Text, NewData) Then
    txtSaveAs.BackColor = RGB(220, 255, 220)
  Else
    txtSaveAs.BackColor = RGB(255, 220, 220)
  End If
End Sub

Private Sub btnClear_Click()
  Erase OldData
  Erase NewData
  txtFlename.Text = vbNullString
  txtFlename.BackColor = RGB(220, 220, 220)
  txtSaveAs.Text = vbNullString
  txtSaveAs.BackColor = RGB(220, 220, 220)
  cboxMethod.Locked = False
  cboxMethod.BackColor = RGB(220, 220, 220)
  cboxMethod.ListIndex = 1
  txtDecompSize.Text = vbNullString
  txtCompSize.Text = vbNullString
  txtPercentage.Text = vbNullString
  btnCompress.Enabled = False
  btnSave.Enabled = False
  cbCompressed.Value = False
  btnClear.Enabled = False
End Sub

Private Sub LoadFile(ByVal fPath As String)
  On Error GoTo Fail

  If Len(fPath) = 0 Then GoTo Fail
  If Len(Dir(fPath)) = 0 Then GoTo Fail

  Erase NewData
  txtDecompSize.Text = vbNullString
  txtCompSize.Text = vbNullString
  txtPercentage.Text = vbNullString
  btnSave.Enabled = False

  OldData = m_Squash.ReadFile(fPath)
  If m_Squash.CheckArray(OldData) = 0 Then GoTo Fail

  txtFlename.Text = fPath
  txtFlename.BackColor = RGB(220, 255, 220)
  btnCompress.Enabled = True

  Dim algo As Long, idx As Long
  algo = m_Squash.IsCompressed(OldData)
  idx = AlgoIndex(algo)

  cbCompressed.Value = (idx >= 0)

  If idx >= 0 Then
    CompressTheFile = False
    btnCompress.Caption = "Decompress"
    btnCompress.Accelerator = "D"
    cboxMethod.ListIndex = idx
    cboxMethod.Locked = True       ' method is fixed by the file's header
    txtCompSize.Text = m_Squash.CheckArray(OldData)
    txtSaveAs.Text = Replace(fPath, ".Compressed", "") & ".Decompressed"
  Else
    CompressTheFile = True
    btnCompress.Caption = "Compress"
    btnCompress.Accelerator = "C"
    cboxMethod.Locked = False
    txtDecompSize.Text = m_Squash.CheckArray(OldData)
    txtSaveAs.Text = fPath & ".Compressed"
  End If

  If Len(Dir(txtSaveAs.Text)) Then
    txtSaveAs.BackColor = RGB(255, 255, 220)   ' will overwrite an existing file
  Else
    txtSaveAs.BackColor = RGB(220, 220, 220)
  End If
  btnClear.Enabled = True
  cboxMethod.BackColor = RGB(220, 255, 220)
  Exit Sub

Fail:
  txtFlename.BackColor = RGB(255, 220, 220)
  btnCompress.Enabled = False
  btnSave.Enabled = False
  Erase OldData
  Erase NewData
End Sub

Private Function AlgoIndex(ByVal algo As Long) As Long
  Dim i As Long
  AlgoIndex = -1
  If algo = 0 Then Exit Function
  For i = LBound(AlgoIds) To UBound(AlgoIds)
    If AlgoIds(i) = algo Then AlgoIndex = i: Exit Function
  Next i
End Function
