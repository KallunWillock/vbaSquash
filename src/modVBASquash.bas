Attribute VB_Name = "modvbaSquash"
'    ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'    |||||||||||||||||           vbaSquash  (v1.0)             ||||||||||||||||||||
'    ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'
'    AUTHOR:   Kallun Willock
'    RESEARCH: MSDN - https://docs.microsoft.com/en-us/windows/win32/api/_cmpapi/
'              Frank Schüler - https://foren.activevb.de/archiv/vb-classic/thread-410382/beitrag-410466/Komprimierung-im-BufferMode/
'              Tanner Helland - https://github.com/tannerhelland/VB6-Compression
'
'    NOTES:    - Uses Win32 APIs in the cabinet.dll library to compress and decompress data.
'              - The API prepends a 12-byte header in Buffer Mode.
'              - Only available on Windows 8+.
'              - Experience with MSZIP has been odd. Apparently its legacy and should
'                avoided if possible. XPRESS is preferred.
'
'    LICENSE:  MIT
'
'    VERSION:  1.0    Uploaded to Github

Option Explicit

Public Enum COMPRESS_ALGORITHM_ENUM
  MSZIP = 2
  XPRESS = 3
  XPRESS_HUFF = 4
  LZMS = 5
End Enum

Private Const INVALID_HANDLE_VALUE As Long = -1&
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const CREATE_ALWAYS As Long = &H2
Private Const OPEN_EXISTING As Long = &H3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const GENERIC_WRITE As Long = &H40000000
Private Const GENERIC_READ As Long = &H80000000
Private Const COMPRESS_RAW As Long = &H20000000
Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122&

#If VBA7 Then
  Private Declare PtrSafe Function CreateCompressor Lib "cabinet.dll" (ByVal Algorithm As Long, ByVal AllocationRoutines As LongPtr, ByRef hCompressor As LongPtr) As Long
  Private Declare PtrSafe Function CompressAPI Lib "cabinet.dll" Alias "Compress" (ByVal hCompressor As LongPtr, ByVal UncompressedData As LongPtr, ByVal UncompressedDataSize As Long, ByVal CompressedBuffer As LongPtr, ByVal CompressedBufferSize As Long, ByRef CompressedDataSizeRequired As Long) As Long
  Private Declare PtrSafe Function CloseCompressor Lib "cabinet.dll" (ByVal hCompressor As LongPtr) As Long
  
  Private Declare PtrSafe Function CreateDecompressor Lib "cabinet.dll" (ByVal Algorithm As Long, ByVal AllocationRoutines As LongPtr, ByRef hDecompressor As LongPtr) As Long
  Private Declare PtrSafe Function DecompressAPI Lib "cabinet.dll" Alias "Decompress" (ByVal hDecompressor As LongPtr, ByVal CompressedData As LongPtr, ByVal CompressedDataSize As Long, ByVal UncompressedBuffer As LongPtr, ByVal UncompressedBufferSize As Long, ByRef UncompressedDataSizeRequired As Long) As Long
  Private Declare PtrSafe Function CloseDecompressor Lib "cabinet.dll" (ByVal hDecompressor As LongPtr) As Long
  
  Private Declare PtrSafe Function apiCreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
  Private Declare PtrSafe Function apiWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToWrite As Long, Optional ByRef lpNumberOfBytesWritten As Long, Optional ByVal lpOverlapped As LongPtr) As Long
  Private Declare PtrSafe Function apiReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToRead As LongLong, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As LongPtr) As Long
  Private Declare PtrSafe Function apiDeleteFile Lib "kernel32" Alias "DeleteFileW" (ByVal lpFileName As LongPtr) As Long
  Private Declare PtrSafe Function apiGetFileSizeEx Lib "kernel32" Alias "GetFileSizeEx" (ByVal hFile As LongPtr, ByRef lpFileSize As LongLong) As Long
  Private Declare PtrSafe Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As LongPtr) As Long
  
  Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
#Else
  ' Dummy for 32-bit to allow code to compile with LongPtr type
  ' Courtesy of https://github.com/OlimilO1402/XL_VBanywhere
  Private Enum LongPtr
    [_]
  End Enum
  
  Private Declare Function CreateCompressor Lib "cabinet.dll" (ByVal Algorithm As Long, ByVal AllocationRoutines As Long, ByRef hCompressor As Long) As Long
  Private Declare Function CompressAPI Lib "cabinet.dll" Alias "Compress" (ByVal hCompressor As Long, ByVal UncompressedData As Long, ByVal UncompressedDataSize As Long, ByVal CompressedBuffer As Long, ByVal CompressedBufferSize As Long, ByRef CompressedDataSizeRequired As Long) As Long
  Private Declare Function CloseCompressor Lib "cabinet.dll" (ByVal hCompressor As Long) As Long
  
  Private Declare Function CreateDecompressor Lib "cabinet.dll" (ByVal Algorithm As Long, ByVal AllocationRoutines As Long, ByRef hDecompressor As Long) As Long
  Private Declare Function DecompressAPI Lib "cabinet.dll" Alias "Decompress" (ByVal hDecompressor As Long, ByVal CompressedData As Long, ByVal CompressedDataSize As Long, ByVal UncompressedBuffer As Long, ByVal UncompressedBufferSize As Long, ByRef UncompressedDataSizeRequired As Long) As Long
  Private Declare Function CloseDecompressor Lib "cabinet.dll" (ByVal hDecompressor As Long) As Long
  
  Private Declare Function apiCreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
  Private Declare Function apiWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
  Private Declare Function apiReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
  Private Declare Function apiDeleteFile Lib "kernel32" Alias "DeleteFileW" (ByVal lpFileName As Long) As Long
  Private Declare Function apiGetFileSizeEx Lib "kernel32" (ByVal hFile As Long, ByRef lpFileSize As Currency) As Long ' Currency for 64-bit integer
  Private Declare Function apiCloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
  
  Private Declare Function GetLastError Lib "kernel32" () As Long
#End If

Public Function CompressBytes(ByRef Source() As Byte, Optional ByVal Algorithm As COMPRESS_ALGORITHM_ENUM = LZMS) As Byte()
  Dim Result As Long, hCompressor As LongPtr, SourceDataSize As Long, RequiredCompressedBufferSize As Long
  Dim ActualBytesWritten As Long, CompressedOutput() As Byte
  
  On Error GoTo ErrHandler
  SourceDataSize = CheckArray(Source)
  If SourceDataSize <= 0 Then Exit Function
  
  Result = CreateCompressor(Algorithm, 0, hCompressor)
  If Result = 0 Or hCompressor = 0 Then Debug.Print "[vbaSquash] CreateCompressor failed: " & GetLastError(): GoTo CleanUp
  
  ' Two-step compression process: Step #1 - First call to CompressAPI with 0 output
  ' buffer to get required size - expected return 0 (FALSE) and GetLastError() should be 122
  Result = CompressAPI(hCompressor, VarPtr(Source(LBound(Source))), SourceDataSize, 0, 0, RequiredCompressedBufferSize)
  
  If Result = 0 And GetLastError() = ERROR_INSUFFICIENT_BUFFER Then
    If RequiredCompressedBufferSize > 0 Then
      ReDim CompressedOutput(0 To RequiredCompressedBufferSize - 1) As Byte
      
      ' Step #2 - Perform actual compression with correctly sized buffer -
      '           non-zero (TRUE) for success
      Result = CompressAPI(hCompressor, VarPtr(Source(LBound(Source))), SourceDataSize, VarPtr(CompressedOutput(0)), RequiredCompressedBufferSize, ActualBytesWritten)
      If Result <> 0 And ActualBytesWritten > 0 Then
        If ActualBytesWritten < RequiredCompressedBufferSize Then
          ReDim Preserve CompressedOutput(0 To ActualBytesWritten - 1)
        End If
        CompressBytes = CompressedOutput
      Else
        Debug.Print "[vbaSquash] Second CompressAPI call failed. LastDllError: " & GetLastError()
      End If
    Else
      ' RequiredCompressedBufferSize was 0 after the query, likely uncompressible or empty output.
      Debug.Print "[vbaSquash] CompressAPI query indicated 0 bytes required for compressed output."
    End If
  Else
    ' A non-true return is not expected to happen.
    Debug.Print "[vbaSquash] First CompressAPI call unexpectedly succeeded. NB: Required Size = " & RequiredCompressedBufferSize
  End If
  
CleanUp:
  If hCompressor <> 0 Then
    CloseCompressor hCompressor
    hCompressor = 0
  End If
  Exit Function
  
ErrHandler:
  Debug.Print "[vbaSquash] Error " & Err.Number & ": " & Err.Description & ". LastDLLError = " & GetLastError()
  GoTo CleanUp
  
End Function

Public Function DecompressBytes(ByRef Source() As Byte, Optional ByVal Algorithm As COMPRESS_ALGORITHM_ENUM = -1) As Byte()
  
  On Error GoTo ErrHandler
  
  If CheckArray(Source) > 0 Then
    Dim Result As Long, hCompressor As LongPtr, BYteLength As Long, DecompressedBufferSize As Long
    If Algorithm = -1 Then Algorithm = IsCompressed(Source)
    If CInt(Algorithm) <= 0 Then Debug.Print "[vbaSquash] Invalid or undetermined algorithm.": GoTo CleanUp
    
    Result = CreateDecompressor(Algorithm, 0, hCompressor)
    
    If Result <> 0 Then
      ReDim Buffer(0) As Byte
      BYteLength = VBA.LenB(Source)
      Result = DecompressAPI(hCompressor, VarPtr(Source(LBound(Source))), BYteLength, 0, 0, DecompressedBufferSize)
      If Result = 0 Then
        ReDim Buffer(0 To DecompressedBufferSize - 1)
        If DecompressAPI(hCompressor, VarPtr(Source(0)), BYteLength, VarPtr(Buffer(0)), DecompressedBufferSize, DecompressedBufferSize) Then
          If DecompressedBufferSize > 0 Then
            ReDim Preserve Buffer(0 To DecompressedBufferSize - 1)
            DecompressBytes = Buffer
          End If
        End If
      End If
      
CleanUp:
      If hCompressor <> 0 Then
        CloseDecompressor hCompressor
        hCompressor = 0
      End If
      Erase Buffer
    End If
  Else
    Debug.Print "[vbaSquash] Source data is empty or invalid."
  End If
  Exit Function
  
ErrHandler:
  Debug.Print "[vbaSquash] Error " & Err.Number & ": " & Err.Description & ". LastDLLError = " & GetLastError()
  GoTo CleanUp
  
End Function

Public Function IsCompressed(ByRef Source As Variant) As COMPRESS_ALGORITHM_ENUM
  Dim Data() As Byte
  On Error GoTo CleanExit
  
  Select Case VarType(Source)
    Case vbString
      Dim TargetFilename As String
      TargetFilename = CStr(Source)
      If Dir(TargetFilename) = "" Or (GetAttr(TargetFilename) And vbDirectory) = vbDirectory Then Exit Function
      Data = ReadFile(TargetFilename, 12)
    Case vbArray + vbByte
      Data = Source
    Case Else
      Exit Function
  End Select
  
  If UBound(Data) >= 7 Then
    ' Check the file's magic number to confirm that it is a compressed file
    If Data(0) = &HA And Data(1) = &H51 And Data(2) = &HE5 And Data(3) = &HC0 And Data(4) = &H18 And Data(5) = &H0 Then
      ' Check the 8th byte for the algorithm
      Select Case Data(7)
        Case MSZIP, XPRESS, XPRESS_HUFF, LZMS
          IsCompressed = Data(7)
        Case Else
          IsCompressed = -1
        End Select
    End If
  End If
  
CleanExit:
  Erase Data
  
End Function

Public Function CompressFile(ByVal TargetFilename As String, Optional ByVal OutputFilename As String = vbNullString, Optional ByVal Algorithm As COMPRESS_ALGORITHM_ENUM = LZMS, Optional ByVal Overwrite As Boolean = False) As Boolean
  Dim FileBytes() As Byte, CompressedData() As Byte, OutFile As String
  
  On Error GoTo ErrHandler
  If Dir(TargetFilename) = "" Or (GetAttr(TargetFilename) And vbDirectory) = vbDirectory Then Exit Function
  
  FileBytes = ReadFile(TargetFilename)
  
  CompressedData = CompressBytes(FileBytes, Algorithm)
  
  OutFile = IIf(LenB(OutputFilename), OutputFilename, TargetFilename & ".compressed")
  CompressFile = WriteFile(OutFile, CompressedData, Overwrite)
  
CleanUp:
  Erase FileBytes
  Erase CompressedData
  Exit Function
  
ErrHandler:
  Debug.Print "[vbaSquash] Error " & Err.Number & ": " & Err.Description & ". LastDLLError = " & GetLastError()
  GoTo CleanUp
  
End Function

Public Function DecompressFile(ByVal TargetFilename As String, Optional ByVal OutputFilename As String = vbNullString, Optional ByVal Algorithm As COMPRESS_ALGORITHM_ENUM = -1, Optional ByVal Overwrite As Boolean = False) As Boolean
  Dim FileBytes() As Byte, DecompressedData() As Byte, OutFile As String
  
  On Error GoTo ErrHandler
  
  If Dir(TargetFilename) = "" Or (GetAttr(TargetFilename) And vbDirectory) = vbDirectory Then Exit Function
  
  FileBytes = ReadFile(TargetFilename)
  
  DecompressedData = DecompressBytes(FileBytes, Algorithm)
  
  OutFile = IIf(LenB(OutputFilename), OutputFilename, TargetFilename & ".decompressed")
  DecompressFile = WriteFile(OutFile, DecompressedData, Overwrite)
  
CleanUp:
  Erase FileBytes
  Erase DecompressedData
  Exit Function
  
ErrHandler:
  Debug.Print "[vbaSquash] Error " & Err.Number & ": " & Err.Description & ". LastDLLError = " & GetLastError()
  GoTo CleanUp
  
End Function

Public Function ReadFile(ByVal TargetFilename As String, Optional ByVal BytesToRead As LongLong = -1, Optional ByVal ExpectString As Boolean = False) As Variant
  Dim hFile As LongPtr, FileSize As LongLong, ActualBytesToRead As LongLong, BytesReadSuccessfully As Long, ResultApi As Long, Buffer() As Byte
  
  On Error GoTo ErrHandler
  
  If Dir(TargetFilename) = "" Or (GetAttr(TargetFilename) And vbDirectory) = vbDirectory Then Exit Function
  
  hFile = apiCreateFile(StrPtr(TargetFilename), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
  If hFile = INVALID_HANDLE_VALUE Then Exit Function
  
  If apiGetFileSizeEx(hFile, FileSize) = 0 Then GoTo CleanUp
  
  ActualBytesToRead = IIf(BytesToRead = -1 Or BytesToRead > FileSize, FileSize, BytesToRead)
  
  If ActualBytesToRead = 0 Or ActualBytesToRead > &H7FFFFFFF Then
    Debug.Print "[vbaSquash] " & IIf(ActualBytesToRead, "File too large to read in VBA safely.", "Nothing to read.")
    ReadFile = IIf(ExpectString, vbNullString, VBA.Array())
    GoTo CleanUp
  End If
   
  ReDim Buffer(0 To CLng(ActualBytesToRead) - 1)
  
  ResultApi = apiReadFile(hFile, VarPtr(Buffer(0)), ActualBytesToRead, BytesReadSuccessfully, 0)
  If ResultApi <> 0 And BytesReadSuccessfully > 0 Then
    If BytesReadSuccessfully < ActualBytesToRead Then ReDim Preserve Buffer(0 To BytesReadSuccessfully - 1)
    ReadFile = IIf(ExpectString, StrConv(Buffer, vbUnicode), Buffer)
  End If
  
CleanUp:
  If hFile <> INVALID_HANDLE_VALUE Then apiCloseHandle hFile
  Exit Function
  
ErrHandler:
  Debug.Print "[vbaSquash] Error " & Err.Number & ": " & Err.Description & ". LastDLLError = " & GetLastError()
  GoTo CleanUp
  
End Function

Public Function WriteFile(ByVal TargetFilename As String, ByRef Data() As Byte, Optional ByVal Overwrite As Boolean = True) As Boolean
  Dim hFile As LongPtr, BytesWritten As Long, DataLength As Long, Result As Long
  
  On Error GoTo ErrHandler
  
  If Not IsArray(Data) Or CheckArray(Data) = 0 Then Exit Function
  DataLength = UBound(Data) - LBound(Data) + 1
  
  If Dir(TargetFilename) <> "" Then
    If (GetAttr(TargetFilename) And vbDirectory) = vbDirectory Then Exit Function
    If Overwrite Then
      If apiDeleteFile(StrPtr(TargetFilename)) = 0 Then Exit Function
    Else
      Exit Function
    End If
  End If
  
  hFile = apiCreateFile(StrPtr(TargetFilename), GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
  If hFile = INVALID_HANDLE_VALUE Then Exit Function
  
  Result = apiWriteFile(hFile, VarPtr(Data(LBound(Data))), DataLength, BytesWritten, 0)
  If Result <> 0 And BytesWritten = DataLength Then
    WriteFile = True
  Else
    If hFile <> INVALID_HANDLE_VALUE Then apiCloseHandle hFile: hFile = INVALID_HANDLE_VALUE
    On Error Resume Next
    apiDeleteFile StrPtr(TargetFilename)
  End If
    
CleanUp:
  If hFile <> INVALID_HANDLE_VALUE Then apiCloseHandle hFile
  Exit Function
  
ErrHandler:
  Debug.Print "[vbaSquash] Error " & Err.Number & ": " & Err.Description & ". LastDLLError = " & GetLastError()
  GoTo CleanUp
  
End Function

Public Function CheckArray(ByRef Source As Variant) As Long

  Dim ub As Long, lb As Long
  On Error Resume Next
  ub = UBound(Source)
  lb = LBound(Source)
  If Err.Number <> 0 Then
    Err.Clear
    Exit Function
  End If
  On Error GoTo 0
  If ub < lb Then Exit Function
  CheckArray = ub - lb + 1
  
End Function

