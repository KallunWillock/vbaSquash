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


Private OldData() As Byte
Private NewData() As Byte
Private CompressTheFile As Boolean
Private CompressionAlgorithm As COMPRESS_ALGORITHM_ENUM
Private IsLoadingFile As Boolean



Private Sub txtFlename_Change()
  If btnClear.Enabled Then btnClear.Enabled = True
End Sub

Private Sub UserForm_Initialize()
  cboxMethod.List = Split("MS-ZIP|XPRESS|XPRESS HUFF|LZMS", "|")
  cboxMethod.ListIndex = 1
  
  btnCompress.Enabled = False
  btnSave.Enabled = False
  btnClear.Enabled = False
  
  txtCompSize.Locked = True
  txtPercentage.Locked = True
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
End Sub

Private Sub btnCompress_Click()
  
  txtCompSize.Text = vbNullString
  txtPercentage.Text = vbNullString
  
  CompressionAlgorithm = cboxMethod.ListIndex + 2
  
  If CompressTheFile Then
    NewData = modvbaSquash.CompressBytes(OldData, CompressionAlgorithm)
  Else
    NewData = modvbaSquash.DecompressBytes(OldData, CompressionAlgorithm)
  End If
  
  Dim originalSize As Long: originalSize = UBound(OldData) + 1
  Dim newSize As Long: newSize = UBound(NewData) + 1
  
  txtDecompSize.Text = IIf(CompressTheFile, originalSize, newSize)
  txtCompSize.Text = IIf(CompressTheFile, newSize, originalSize)
  txtPercentage.Text = Format(100 * (newSize / originalSize), "0.0") & "%"
  
  btnSave.Enabled = True
  
End Sub

Private Sub btnSave_Click()
  On Error Resume Next
  modvbaSquash.WriteFile txtSaveAs.Text, NewData
  txtSaveAs.BackColor = RGB(220, 255, 220)
End Sub

Private Sub btnClear_Click()
  Erase OldData
  Erase NewData
  txtFlename.Text = vbNullString
  txtFlename.BackColor = RGB(220, 220, 220)
  txtSaveAs.Text = vbNullString
  txtSaveAs.BackColor = RGB(220, 220, 220)
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
  
  If Len(Dir(fPath)) = 0 Then GoTo Fail
  
  txtFlename.Text = fPath
  txtFlename.BackColor = RGB(220, 255, 220)
  btnCompress.Enabled = True
  
  OldData = modvbaSquash.ReadFile(fPath)
  
  Dim algo As Long
  algo = modvbaSquash.IsCompressed(OldData)
  
  cbCompressed.Value = (algo <> 0)
  
  If algo Then
    CompressTheFile = False
    btnCompress.Caption = "Decompress"
    btnCompress.Accelerator = "D"
    cboxMethod.ListIndex = algo - 2
    txtCompSize.Text = UBound(OldData) + 1
    txtSaveAs.Text = Replace(fPath, ".Compressed", "") & ".Decompressed"
  Else
    CompressTheFile = True
    btnCompress.Enabled = True
    btnCompress.Caption = "Compress"
    btnCompress.Accelerator = "CD"
    txtDecompSize.Text = UBound(OldData) + 1
    txtSaveAs.Text = fPath & ".Compressed"
  End If
  
  If Len(Dir(txtSaveAs.Text)) Then
    txtSaveAs.BackColor = RGB(255, 255, 220)
  End If
  btnClear.Enabled = True
  cboxMethod.BackColor = RGB(220, 255, 220)
  
  Exit Sub
  
Fail:
  txtFlename.BackColor = RGB(255, 220, 220)
  btnCompress.Enabled = False
  Erase OldData
  Erase NewData
End Sub

