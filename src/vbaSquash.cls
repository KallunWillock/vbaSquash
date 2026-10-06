VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbaSquash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'    |||||||||||||||||           vbaSquash  (v1.1)             ||||||||||||||||||||
'    ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'
'    AUTHOR:   Kallun Willock
'    RESEARCH: MSDN - https://docs.microsoft.com/en-us/windows/win32/api/_cmpapi/
'              MSDN - https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/ntifs/nf-ntifs-rtlcompressbuffer
'              Frank Schuler - https://foren.activevb.de/archiv/vb-classic/thread-410382/beitrag-410466/Komprimierung-im-BufferMode/
'              Tanner Helland - https://github.com/tannerhelland/VB6-Compression
'
'    NOTES:    - Uses Win32 system APIs in the cabinet.dll and ntdll.dll libraries to
'                compress and decompress data.
'
'              - CABINET.DLL  (Algorithm values 2-5):
'                  The API prepends a 12-byte header in Buffer Mode.
'                  Only available on Windows 8+.
'                  Experience with MSZIP has been odd. Apparently it is legacy and
'                  should be avoided if possible. XPRESS is preferred.
'
'              - NTDLL.DLL  (Algorithm values 12-14):
'                  RtlCompressBuffer does not prepend any header.
'                  Because the original size must be known at decompression time,
'                  a custom 12-byte vbaSquash header is optionally prepended.
'                  Header layout:
'                    Bytes  0-5  : Magic "SQUASH"
'                    Byte   6    : Encoded engine  (0=Standard, 2=Maximum)
'                                  Encoded as Engine ^ (1/8); decoded as Data(6) ^ 8
'                    Byte   7    : Algorithm enum value (12/13/14)
'                    Bytes  8-11 : Uncompressed size (Little-endian Long)
'
'              - Routing: Algorithm < 10  -> cabinet.dll
'                         Algorithm >= 10 -> ntdll.dll
'                         Internal ntdll Format = Algorithm - 10
'
'    LICENSE:  MIT
'
'    VERSION:  1.0    07/05/2025  Uploaded to Github
'              1.1    06/10/2026  Extended to ntdll RtlCompressor; unified enum;
'                                 added optional vbaSquash header for Rtl path.

Option Explicit

Public Enum COMPRESS_ALGORITHM_ENUM
  ' cabinet.dll  (Algorithm < 10)
  MSZIP = 2
  XPRESS = 3
  XPRESS_HUFF = 4
  LZMS = 5
  ' ntdll.dll  (Algorithm >= 10. Internal Format = Algorithm - 10)
  RTL_LZNT1 = 12          ' maps to COMPRESSION_FORMAT_LZNT1 = 2
  RTL_XPRESS = 13         ' maps to COMPRESSION_FORMAT_XPRESS = 3
  RTL_XPRESS_HUFFMAN = 14 ' maps to COMPRESSION_FORMAT_XPRESS_HUFFMAN = 4
End Enum

Public Enum NtCompressionEngine
  COMPRESSION_ENGINE_STANDARD = 0
  COMPRESSION_ENGINE_MAXIMUM = 256
End Enum

Private Type LARGE_INTEGER
  LowPart  As Long
  HighPart As Long
End Type

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

Private Const STATUS_SUCCESS As Long = 0
Private Const STATUS_INVALID_BUFFER_SIZE As Long = &HC0000242
Private Const STATUS_BUFFER_TOO_SMALL As Long = &HC0000023
Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
Private Const MEM_RELEASE As Long = &H8000
Private Const PAGE_READWRITE As Long = &H4

Private Const SQUASH_VERSION As Single = 1.1

#If VBA7 Then
  ' cabinet.dll
  Private Declare PtrSafe Function CreateCompressor Lib "cabinet" (ByVal Algorithm As Long, ByVal AllocationRoutines As LongPtr, ByRef hCompressor As LongPtr) As Long
  Private Declare PtrSafe Function CompressAPI Lib "cabinet" Alias "Compress" (ByVal hCompressor As LongPtr, ByVal UncompressedData As LongPtr, ByVal UncompressedDataSize As LongPtr, ByVal CompressedBuffer As LongPtr, ByVal CompressedBufferSize As LongPtr, ByRef CompressedDataSizeRequired As LongPtr) As Long
  Private Declare PtrSafe Function CloseCompressor Lib "cabinet" (ByVal hCompressor As LongPtr) As Long
  Private Declare PtrSafe Function CreateDecompressor Lib "cabinet" (ByVal Algorithm As Long, ByVal AllocationRoutines As LongPtr, ByRef hDecompressor As LongPtr) As Long
  Private Declare PtrSafe Function DecompressAPI Lib "cabinet" Alias "Decompress" (ByVal hDecompressor As LongPtr, ByVal CompressedData As LongPtr, ByVal CompressedDataSize As LongPtr, ByVal UncompressedBuffer As LongPtr, ByVal UncompressedBufferSize As LongPtr, ByRef UncompressedDataSizeRequired As LongPtr) As Long
  Private Declare PtrSafe Function CloseDecompressor Lib "cabinet" (ByVal hDecompressor As LongPtr) As Long

  ' ntdll.dll
  Private Declare PtrSafe Function RtlGetCompressionWorkSpaceSize Lib "ntdll" (ByVal CompressionFormatAndEngine As Long, ByRef CompressBufferWorkSpaceSize As Long, ByRef CompressFragmentWorkSpaceSize As Long) As Long
  Private Declare PtrSafe Function NtAllocateVirtualMemory Lib "ntdll" (ByVal ProcessHandle As LongPtr, ByRef BaseAddress As LongPtr, ByVal ZeroBits As LongPtr, ByRef RegionSize As LongPtr, ByVal AllocationType As Long, ByVal Protect As Long) As Long
  Private Declare PtrSafe Function NtFreeVirtualMemory Lib "ntdll" (ByVal ProcessHandle As LongPtr, ByRef BaseAddress As LongPtr, ByRef RegionSize As LongPtr, ByVal FreeType As Long) As Long
  Private Declare PtrSafe Function RtlCompressBuffer Lib "ntdll" (ByVal CompressionFormatAndEngine As Long, ByVal UncompressedBuffer As LongPtr, ByVal UncompressedBufferSize As Long, ByVal CompressedBuffer As LongPtr, ByVal CompressedBufferSize As Long, ByVal UncompressedChunkSize As Long, ByRef FinalCompressedSize As Long, ByVal WorkSpace As LongPtr) As Long
  Private Declare PtrSafe Function RtlDecompressBuffer Lib "ntdll" (ByVal CompressionFormat As Long, ByVal UncompressedBuffer As LongPtr, ByVal UncompressedBufferSize As Long, ByVal CompressedBuffer As LongPtr, ByVal CompressedBufferSize As Long, ByRef FinalUncompressedSize As Long) As Long
  Private Declare PtrSafe Function RtlDecompressBufferEx Lib "ntdll" (ByVal CompressionFormat As Long, ByVal UncompressedBuffer As LongPtr, ByVal UncompressedBufferSize As Long, ByVal CompressedBuffer As LongPtr, ByVal CompressedBufferSize As Long, ByRef FinalUncompressedSize As Long, ByVal WorkSpace As LongPtr) As Long
  
  ' kernel32
  Private Declare PtrSafe Function apiCreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
  Private Declare PtrSafe Function apiWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToWrite As Long, Optional ByRef lpNumberOfBytesWritten As Long, Optional ByVal lpOverlapped As LongPtr) As Long
  Private Declare PtrSafe Function apiReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As LongPtr) As Long
  Private Declare PtrSafe Function apiDeleteFile Lib "kernel32" Alias "DeleteFileW" (ByVal lpFileName As LongPtr) As Long
  Private Declare PtrSafe Function apiGetFileSizeEx Lib "kernel32" Alias "GetFileSizeEx" (ByVal hFile As LongPtr, ByRef lpFileSize As LARGE_INTEGER) As Long
  Private Declare PtrSafe Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As LongPtr) As Long
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)

#Else
  ' Dummy for 32-bit to allow code to compile with LongPtr type
  ' Courtesy of https://github.com/OlimilO1402/XL_VBanywhere
  Private Enum LongPtr
    [_]
  End Enum

  ' cabinet.dll
  Private Declare Function CreateCompressor Lib "cabinet" (ByVal Algorithm As Long, ByVal AllocationRoutines As Long, ByRef hCompressor As Long) As Long
  Private Declare Function CompressAPI Lib "cabinet" Alias "Compress" (ByVal hCompressor As Long, ByVal UncompressedData As Long, ByVal UncompressedDataSize As Long, ByVal CompressedBuffer As Long, ByVal CompressedBufferSize As Long, ByRef CompressedDataSizeRequired As Long) As Long
  Private Declare Function CloseCompressor Lib "cabinet" (ByVal hCompressor As Long) As Long
  Private Declare Function CreateDecompressor Lib "cabinet" (ByVal Algorithm As Long, ByVal AllocationRoutines As Long, ByRef hDecompressor As Long) As Long
  Private Declare Function DecompressAPI Lib "cabinet" Alias "Decompress" (ByVal hDecompressor As Long, ByVal CompressedData As Long, ByVal CompressedDataSize As Long, ByVal UncompressedBuffer As Long, ByVal UncompressedBufferSize As Long, ByRef UncompressedDataSizeRequired As Long) As Long
  Private Declare Function CloseDecompressor Lib "cabinet" (ByVal hDecompressor As Long) As Long

  ' ntdll.dll
  Private Declare Function RtlGetCompressionWorkSpaceSize Lib "ntdll" (ByVal CompressionFormatAndEngine As Long, ByRef CompressBufferWorkSpaceSize As Long, ByRef CompressFragmentWorkSpaceSize As Long) As Long
  Private Declare Function NtAllocateVirtualMemory Lib "ntdll" (ByVal ProcessHandle As Long, ByRef BaseAddress As Long, ByVal ZeroBits As Long, ByRef RegionSize As Long, ByVal AllocationType As Long, ByVal Protect As Long) As Long
  Private Declare Function NtFreeVirtualMemory Lib "ntdll" (ByVal ProcessHandle As Long, ByRef BaseAddress As Long, ByRef RegionSize As Long, ByVal FreeType As Long) As Long
  Private Declare Function RtlCompressBuffer Lib "ntdll" (ByVal CompressionFormatAndEngine As Long, ByVal UncompressedBuffer As Long, ByVal UncompressedBufferSize As Long, ByVal CompressedBuffer As Long, ByVal CompressedBufferSize As Long, ByVal UncompressedChunkSize As Long, ByRef FinalCompressedSize As Long, ByVal WorkSpace As Long) As Long
  Private Declare Function RtlDecompressBuffer Lib "ntdll" (ByVal CompressionFormat As Long, ByVal UncompressedBuffer As Long, ByVal UncompressedBufferSize As Long, ByVal CompressedBuffer As Long, ByVal CompressedBufferSize As Long, ByRef FinalUncompressedSize As Long) As Long

  ' kernel32
  Private Declare Function apiCreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
  Private Declare Function apiWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
  Private Declare Function apiReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
  Private Declare Function apiDeleteFile Lib "kernel32" Alias "DeleteFileW" (ByVal lpFileName As Long) As Long
  Private Declare Function apiGetFileSizeEx Lib "kernel32" Alias "GetFileSizeEx" (ByVal hFile As Long, ByRef lpFileSize As LARGE_INTEGER) As Long
  Private Declare Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If

' Compresses a byte array using either cabinet.dll (Algorithm 2-5) or
' ntdll.dll RtlCompressBuffer (Algorithm 12-14).
' For the Rtl path, UseHeader prepends a 12-byte vbaSquash header so that
' DecompressBytes can auto-detect the algorithm and original size.
Public Function CompressBytes(ByRef Source() As Byte, _
                              Optional ByVal Algorithm As COMPRESS_ALGORITHM_ENUM = LZMS, _
                              Optional ByVal Engine As NtCompressionEngine = COMPRESSION_ENGINE_STANDARD, _
                              Optional ByVal UseHeader As Boolean = True) As Byte()
  On Error GoTo ErrHandler
  If CheckArray(Source) = 0 Then Exit Function

  If Algorithm < 10 Then
    CompressBytes = CompressCabinet(Source, Algorithm)
  Else
    CompressBytes = CompressRtl(Source, Algorithm, Engine, UseHeader)
  End If

  Exit Function
ErrHandler:
  Debug.Print "[vbaSquash] CompressBytes error " & Err.Number & ": " & Err.Description
End Function

' Decompresses a byte array produced by CompressBytes.
' For cabinet data the algorithm is auto-detected from the cabinet header.
' For Rtl data the algorithm and original size are read from the vbaSquash
' header if present; otherwise Algorithm and OriginalSize must be supplied.
Public Function DecompressBytes(ByRef Source() As Byte, _
                                Optional ByVal Algorithm As COMPRESS_ALGORITHM_ENUM = -1, _
                                Optional ByVal Engine As NtCompressionEngine = COMPRESSION_ENGINE_STANDARD, _
                                Optional ByVal OriginalSize As Long = 0) As Byte()
  On Error GoTo ErrHandler

  If CheckArray(Source) = 0 Then
    Debug.Print "[vbaSquash] Source data is empty or invalid."
    Exit Function
  End If

  If Algorithm = -1 Then Algorithm = IsCompressed(Source)
  If CInt(Algorithm) <= 0 Then
    Debug.Print "[vbaSquash] Invalid or undetermined algorithm."
    Exit Function
  End If

  If Algorithm < 10 Then
    DecompressBytes = DecompressCabinet(Source, Algorithm)
  Else
    DecompressBytes = DecompressRtl(Source, Algorithm, Engine, OriginalSize)
  End If

  Exit Function
ErrHandler:
  Debug.Print "[vbaSquash] DecompressBytes error " & Err.Number & ": " & Err.Description
End Function

' Returns the COMPRESS_ALGORITHM_ENUM value if the data carries a recognised header,
' or 0 if unrecognised. Source may be a byte array or a file path string.
Public Function IsCompressed(ByRef Source As Variant) As COMPRESS_ALGORITHM_ENUM
  Dim data() As Byte
  On Error GoTo CleanExit

  Select Case VarType(Source)
    Case vbString
      Dim TargetFilename As String
      TargetFilename = CStr(Source)
      If Dir(TargetFilename) = "" Or (GetAttr(TargetFilename) And vbDirectory) = vbDirectory Then Exit Function
      data = ReadFile(TargetFilename, 12)
    Case vbArray + vbByte
      data = Source
    Case Else
      Exit Function
  End Select

  If UBound(data) < 7 Then GoTo CleanExit

  ' cabinet.dll buffer-mode magic: 0A 51 E5 C0 18 00
  If data(0) = &HA And data(1) = &H51 And data(2) = &HE5 And _
     data(3) = &HC0 And data(4) = &H18 And data(5) = &H0 Then
    Select Case data(7)
      Case MSZIP, XPRESS, XPRESS_HUFF, LZMS
        IsCompressed = data(7)
    End Select
    GoTo CleanExit
  End If

  ' vbaSquash header magic: "SQUASH"
  If data(0) = &H53 And data(1) = &H51 And data(2) = &H55 And _
     data(3) = &H41 And data(4) = &H53 And data(5) = &H48 Then
    Select Case data(7)
      Case RTL_LZNT1, RTL_XPRESS, RTL_XPRESS_HUFFMAN
        IsCompressed = data(7)
    End Select
  End If

CleanExit:
  Erase data
End Function

Public Function CompressFile(ByVal TargetFilename As String, _
                             Optional ByVal OutputFilename As String = vbNullString, _
                             Optional ByVal Algorithm As COMPRESS_ALGORITHM_ENUM = LZMS, _
                             Optional ByVal Engine As NtCompressionEngine = COMPRESSION_ENGINE_STANDARD, _
                             Optional ByVal Overwrite As Boolean = False) As Boolean
  Dim FileBytes() As Byte, CompressedData() As Byte, OutFile As String

  On Error GoTo ErrHandler
  If Dir(TargetFilename) = "" Or (GetAttr(TargetFilename) And vbDirectory) = vbDirectory Then Exit Function

  FileBytes = ReadFile(TargetFilename)
  CompressedData = CompressBytes(FileBytes, Algorithm, Engine)

  OutFile = IIf(LenB(OutputFilename), OutputFilename, TargetFilename)
  CompressFile = WriteFile(OutFile, CompressedData, Overwrite)

CleanUp:
  Erase FileBytes: Erase CompressedData
  Exit Function
ErrHandler:
  Debug.Print "[vbaSquash] CompressFile error " & Err.Number & ": " & Err.Description & ". LastDLLError = " & Err.LastDllError
  GoTo CleanUp
End Function

Public Function DecompressFile(ByVal TargetFilename As String, _
                               Optional ByVal OutputFilename As String = vbNullString, _
                               Optional ByVal Algorithm As COMPRESS_ALGORITHM_ENUM = -1, _
                               Optional ByVal Overwrite As Boolean = False) As Boolean
  Dim FileBytes() As Byte, DecompressedData() As Byte, OutFile As String

  On Error GoTo ErrHandler
  If Dir(TargetFilename) = "" Or (GetAttr(TargetFilename) And vbDirectory) = vbDirectory Then Exit Function

  FileBytes = ReadFile(TargetFilename)
  DecompressedData = DecompressBytes(FileBytes, Algorithm)

  OutFile = IIf(LenB(OutputFilename), OutputFilename, TargetFilename)
  DecompressFile = WriteFile(OutFile, DecompressedData, Overwrite)

CleanUp:
  Erase FileBytes: Erase DecompressedData
  Exit Function
ErrHandler:
  Debug.Print "[vbaSquash] DecompressFile error " & Err.Number & ": " & Err.Description & ". LastDLLError = " & Err.LastDllError
  GoTo CleanUp
End Function

Public Function ReadFile(ByVal TargetFilename As String, _
                         Optional ByVal BytesToRead As Long = -1, _
                         Optional ByVal ExpectString As Boolean = False) As Variant
  Dim hFile As LongPtr, ActualBytesToRead As Long
  Dim BytesReadSuccessfully As Long, ResultApi As Long, Buffer() As Byte
  Dim FileSize As LARGE_INTEGER

  On Error GoTo ErrHandler

  If Len(TargetFilename) = 0 Then Exit Function
  If Dir(TargetFilename) = "" Or (GetAttr(TargetFilename) And vbDirectory) = vbDirectory Then Exit Function

  hFile = apiCreateFile(StrPtr(TargetFilename), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
  If hFile = INVALID_HANDLE_VALUE Then Exit Function

  If apiGetFileSizeEx(hFile, FileSize) = 0 Then GoTo CleanUp

    ' Files of 2 GB or more cannot be held in a VBA array.
  If FileSize.HighPart <> 0 Or FileSize.LowPart < 0 Then
    Debug.Print "[vbaSquash] File too large to read in VBA safely."
    ReadFile = IIf(ExpectString, vbNullString, VBA.Array())
    GoTo CleanUp
  End If

  ActualBytesToRead = FileSize.LowPart
  If BytesToRead >= 0 And BytesToRead < ActualBytesToRead Then ActualBytesToRead = BytesToRead

  If ActualBytesToRead = 0 Then
    Debug.Print "[vbaSquash] Nothing to read."
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
  Debug.Print "[vbaSquash] ReadFile error " & Err.Number & ": " & Err.Description & ". LastDLLError = " & Err.LastDllError
  GoTo CleanUp
End Function

Public Function WriteFile(ByVal TargetFilename As String, _
                          ByRef data() As Byte, _
                          Optional ByVal Overwrite As Boolean = True) As Boolean
  Dim hFile As LongPtr, BytesWritten As Long, DataLength As Long, result As Long

  On Error GoTo ErrHandler

  DataLength = CheckArray(data)
  If DataLength = 0 Then Exit Function

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

  result = apiWriteFile(hFile, VarPtr(data(LBound(data))), DataLength, BytesWritten, 0)
  If result <> 0 And BytesWritten = DataLength Then
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
  Debug.Print "[vbaSquash] WriteFile error " & Err.Number & ": " & Err.Description & ". LastDLLError = " & Err.LastDllError
  GoTo CleanUp
End Function

Public Function CheckArray(ByRef Source As Variant) As Long
  If Not IsArray(Source) Then Exit Function
  Dim ub As Long, lb As Long
  On Error Resume Next
  ub = UBound(Source): lb = LBound(Source)
  If Err.Number <> 0 Then Err.Clear: Exit Function
  On Error GoTo 0
  If ub < lb Then Exit Function
  CheckArray = ub - lb + 1
End Function

' =============
'  CABINET.DLL
' =============

Private Function CompressCabinet(ByRef Source() As Byte, _
                                 ByVal Algorithm As COMPRESS_ALGORITHM_ENUM) As Byte()
  Dim result As Long, hCompressor As LongPtr
  Dim SourceDataSize As Long, RequiredSize As LongPtr, ActualBytesWritten As LongPtr
  Dim CompressedOutput() As Byte

  On Error GoTo ErrHandler

  SourceDataSize = CheckArray(Source)

  result = CreateCompressor(Algorithm, 0, hCompressor)
  If result = 0 Or hCompressor = 0 Then
    Debug.Print "[vbaSquash] CreateCompressor failed: " & Err.LastDllError
    GoTo CleanUp
  End If

  ' Step 1: query required output size. With a null output buffer the call cannot
  ' genuinely succeed, so the reported RequiredSize is the real signal. The error
  ' code (ERROR_INSUFFICIENT_BUFFER) is read from Err.LastDllError, never from the
  ' GetLastError API, because apparently the VBA runtime can overwrite the thread's last-error
  ' value between the Declare call returning and the next line running. Or so I'm told.
  result = CompressAPI(hCompressor, VarPtr(Source(LBound(Source))), SourceDataSize, 0, 0, RequiredSize)
  Dim ProbeError As Long: ProbeError = Err.LastDllError

  If result = 0 And RequiredSize > 0 Then
    If RequiredSize > 0 Then
      ReDim CompressedOutput(0 To CLng(RequiredSize) - 1)

      ' Step 2: perform compression
      result = CompressAPI(hCompressor, VarPtr(Source(LBound(Source))), SourceDataSize, VarPtr(CompressedOutput(0)), RequiredSize, ActualBytesWritten)
      If result <> 0 And ActualBytesWritten > 0 Then
        If ActualBytesWritten < RequiredSize Then ReDim Preserve CompressedOutput(0 To CLng(ActualBytesWritten) - 1)
        CompressCabinet = CompressedOutput
      Else
        Debug.Print "[vbaSquash] CompressAPI (step 2) failed. LastDllError: " & Err.LastDllError
      End If
    Else
      Debug.Print "[vbaSquash] CompressAPI query returned 0 bytes required."
    End If
  Else
    Debug.Print "[vbaSquash] CompressAPI (step 1) size query failed. Returned " & result & ", required size " & RequiredSize & ", LastDllError " & ProbeError
  End If

CleanUp:
  If hCompressor <> 0 Then CloseCompressor hCompressor
  Exit Function
ErrHandler:
  Debug.Print "[vbaSquash] CompressCabinet error " & Err.Number & ": " & Err.Description & ". LastDLLError = " & Err.LastDllError
  GoTo CleanUp
End Function

Private Function DecompressCabinet(ByRef Source() As Byte, _
                                   ByVal Algorithm As COMPRESS_ALGORITHM_ENUM) As Byte()
  Dim result As Long, hDecompressor As LongPtr
  Dim ByteLength As Long, DecompressedBufferSize As LongPtr
  Dim Buffer() As Byte

  On Error GoTo ErrHandler

  ByteLength = CheckArray(Source)

  result = CreateDecompressor(Algorithm, 0, hDecompressor)
  If result = 0 Or hDecompressor = 0 Then
    Debug.Print "[vbaSquash] CreateDecompressor failed: " & Err.LastDllError
    GoTo CleanUp
  End If

  ' Step 1: query required output size
  result = DecompressAPI(hDecompressor, VarPtr(Source(LBound(Source))), ByteLength, 0, 0, DecompressedBufferSize)
  If result = 0 And DecompressedBufferSize > 0 Then
    ReDim Buffer(0 To CLng(DecompressedBufferSize) - 1)

    ' Step 2: perform decompression
    If DecompressAPI(hDecompressor, VarPtr(Source(LBound(Source))), ByteLength, VarPtr(Buffer(0)), DecompressedBufferSize, DecompressedBufferSize) Then
      If DecompressedBufferSize > 0 Then
        ReDim Preserve Buffer(0 To CLng(DecompressedBufferSize) - 1)
        DecompressCabinet = Buffer
      End If
    Else
      Debug.Print "[vbaSquash] DecompressAPI (step 2) failed. LastDllError: " & Err.LastDllError
    End If
  Else
    Debug.Print "[vbaSquash] DecompressAPI (step 1) failed. LastDllError: " & Err.LastDllError
  End If

CleanUp:
  If hDecompressor <> 0 Then CloseDecompressor hDecompressor
  Erase Buffer
  Exit Function
ErrHandler:
  Debug.Print "[vbaSquash] DecompressCabinet error " & Err.Number & ": " & Err.Description & ". LastDLLError = " & Err.LastDllError
  GoTo CleanUp
End Function

' ===========
'  NTDLL.DLL
' ===========

Private Function CompressRtl(ByRef Source() As Byte, _
                              ByVal Algorithm As COMPRESS_ALGORITHM_ENUM, _
                              ByVal Engine As NtCompressionEngine, _
                              ByVal UseHeader As Boolean) As Byte()
  Dim InputSize As Long, OutputBufferSize As Long, FinalSize As Long
  Dim WorkSpaceSize As Long, FragmentWorkSpaceSize As Long
  Dim WorkSpace As LongPtr, RegionSize As LongPtr
  Dim FormatAndEngine As Long, Status As Long
  Dim RawCompressed() As Byte

  On Error GoTo ErrHandler

  If Algorithm < RTL_LZNT1 Or Algorithm > RTL_XPRESS_HUFFMAN Then
    Debug.Print "[vbaSquash] CompressRtl: unsupported algorithm " & Algorithm
    Exit Function
  End If
  InputSize = CheckArray(Source)
  If InputSize = 0 Then Exit Function
  FormatAndEngine = (Algorithm - 10) Or CLng(Engine)   ' internal Format = Algorithm - 10

  ' 1. Get workspace size
  Status = RtlGetCompressionWorkSpaceSize(FormatAndEngine, WorkSpaceSize, FragmentWorkSpaceSize)
  If Status <> STATUS_SUCCESS Then
    Debug.Print "[vbaSquash] RtlGetCompressionWorkSpaceSize failed: &H" & Hex(Status)
    GoTo CleanUp
  End If

  ' 2. Allocate workspace
  RegionSize = WorkSpaceSize
  Status = NtAllocateVirtualMemory(-1, WorkSpace, 0, RegionSize, MEM_COMMIT Or MEM_RESERVE, PAGE_READWRITE)
  If Status <> STATUS_SUCCESS Then
    Debug.Print "[vbaSquash] NtAllocateVirtualMemory failed: &H" & Hex(Status)
    WorkSpace = 0
    GoTo CleanUp
  End If

  ' 3. Allocate output buffer. LZNT1 can expand incompressible data, so allow
  '    a generous margin rather than retrying on STATUS_BUFFER_TOO_SMALL.
  OutputBufferSize = InputSize + (InputSize \ 8) + 4096
  ReDim RawCompressed(0 To OutputBufferSize - 1)

  ' 4. Compress
  Status = RtlCompressBuffer(FormatAndEngine, VarPtr(Source(LBound(Source))), InputSize, VarPtr(RawCompressed(0)), OutputBufferSize, 4096, FinalSize, WorkSpace)
  ' NTSTATUS values with the top bit clear are successes. In particular
  ' STATUS_BUFFER_ALL_ZEROS (&H117) means "compressed, and the input was all zeros".
  If Status < 0 Then
    If Status = STATUS_BUFFER_TOO_SMALL And InputSize < 16 Then
      Debug.Print "[vbaSquash] RtlCompressBuffer rejected a " & InputSize & "-byte input (too small for this format)."
    Else
      Debug.Print "[vbaSquash] RtlCompressBuffer failed: &H" & Hex(Status)
    End If
    GoTo CleanUp
  End If
  If FinalSize <= 0 Then GoTo CleanUp
  ReDim Preserve RawCompressed(0 To FinalSize - 1)

  ' 5. Optionally prepend vbaSquash header
  If UseHeader Then
    CompressRtl = BuildSquashHeader(InputSize, Algorithm, Engine, RawCompressed)
  Else
    CompressRtl = RawCompressed
  End If

CleanUp:
  ' RegionSize must be 0 for MEM_RELEASE
  If WorkSpace <> 0 Then
    RegionSize = 0
    NtFreeVirtualMemory -1, WorkSpace, RegionSize, MEM_RELEASE
    WorkSpace = 0
  End If
  Exit Function
ErrHandler:
  Debug.Print "[vbaSquash] CompressRtl error " & Err.Number & ": " & Err.Description
  Resume CleanUp
End Function

Private Function DecompressRtl(ByRef Source() As Byte, _
                                ByVal Algorithm As COMPRESS_ALGORITHM_ENUM, _
                                ByVal Engine As NtCompressionEngine, _
                                ByVal OriginalSize As Long) As Byte()
  Dim Status As Long, FinalSize As Long
  Dim Payload() As Byte, Dest() As Byte
  Dim lb As Long, SourceSize As Long, PayloadOffset As Long, PayloadSize As Long
  Dim WorkSpaceSize As Long, FragmentWorkSpaceSize As Long
  Dim WorkSpace As LongPtr, RegionSize As LongPtr

  On Error GoTo ErrHandler

  SourceSize = CheckArray(Source)
  If SourceSize = 0 Then Exit Function
  lb = LBound(Source)

  ' Read engine, algorithm and original size from the vbaSquash header if present.
  ' Only inspect the header when there are at least 12 bytes to read.
  If SourceSize >= 12 Then
    If Source(lb) = &H53 And Source(lb + 1) = &H51 And Source(lb + 2) = &H55 And _
       Source(lb + 3) = &H41 And Source(lb + 4) = &H53 And Source(lb + 5) = &H48 Then
      Engine = CLng(Source(lb + 6)) ^ 8   ' decode: 0->0, 2->256 (informational only)
      Algorithm = Source(lb + 7)          ' 12, 13, or 14
      CopyMemory VarPtr(OriginalSize), VarPtr(Source(lb + 8)), 4
      PayloadOffset = 12
    End If
  End If

  ' Validate before anything reaches ntdll: a corrupted header byte 7 must not
  ' become an arbitrary compression format.
  If Algorithm < RTL_LZNT1 Or Algorithm > RTL_XPRESS_HUFFMAN Then
    Debug.Print "[vbaSquash] DecompressRtl: unsupported or undetermined algorithm " & Algorithm
    Exit Function
  End If
  If OriginalSize <= 0 Then
    Debug.Print "[vbaSquash] OriginalSize must be provided when data has no vbaSquash header."
    Exit Function
  End If

  ' Extract compressed payload (skip header if present)
  PayloadSize = SourceSize - PayloadOffset
  If PayloadSize <= 0 Then Exit Function
  ReDim Payload(0 To PayloadSize - 1)
  CopyMemory VarPtr(Payload(0)), VarPtr(Source(lb + PayloadOffset)), PayloadSize

  ReDim Dest(0 To OriginalSize - 1)

  ' RtlDecompressBufferEx works for all three Rtl formats. LZNT1 and XPRESS return
  ' a FragmentWorkSpaceSize of 0 (no workspace needed); XPRESS_HUFFMAN requires one.
  ' NtAllocateVirtualMemory rejects a size of 0, so we only allocate when needed.
  Status = RtlGetCompressionWorkSpaceSize(Algorithm - 10, WorkSpaceSize, FragmentWorkSpaceSize)
  If Status <> STATUS_SUCCESS Then
    Debug.Print "[vbaSquash] RtlGetCompressionWorkSpaceSize (decomp) failed: &H" & Hex(Status)
    GoTo CleanUp
  End If
  If FragmentWorkSpaceSize > 0 Then
    RegionSize = FragmentWorkSpaceSize
    Status = NtAllocateVirtualMemory(-1, WorkSpace, 0, RegionSize, MEM_COMMIT Or MEM_RESERVE, PAGE_READWRITE)
    If Status <> STATUS_SUCCESS Then
      Debug.Print "[vbaSquash] NtAllocateVirtualMemory (decomp) failed: &H" & Hex(Status)
      WorkSpace = 0
      GoTo CleanUp
    End If
  End If

  Status = RtlDecompressBufferEx(Algorithm - 10, VarPtr(Dest(0)), OriginalSize, VarPtr(Payload(0)), PayloadSize, FinalSize, WorkSpace)
  If Status <> STATUS_SUCCESS Then
    If Status = STATUS_INVALID_BUFFER_SIZE Then
      Debug.Print "[vbaSquash] RtlDecompressBufferEx: STATUS_INVALID_BUFFER_SIZE - check format and OriginalSize."
    Else
      Debug.Print "[vbaSquash] RtlDecompressBufferEx failed: &H" & Hex(Status)
    End If
    GoTo CleanUp
  End If

  If FinalSize > 0 Then
    If FinalSize < OriginalSize Then ReDim Preserve Dest(0 To FinalSize - 1)
    DecompressRtl = Dest
  End If

CleanUp:
  If WorkSpace <> 0 Then
    RegionSize = 0
    NtFreeVirtualMemory -1, WorkSpace, RegionSize, MEM_RELEASE
    WorkSpace = 0
  End If
  Exit Function
ErrHandler:
  Debug.Print "[vbaSquash] DecompressRtl error " & Err.Number & ": " & Err.Description
  Resume CleanUp
End Function

' Builds a 12-byte vbaSquash header prepended to RawData and returns the combined array.
Private Function BuildSquashHeader(ByVal UncompressedSize As Long, _
                                   ByVal Algorithm As COMPRESS_ALGORITHM_ENUM, _
                                   ByVal Engine As NtCompressionEngine, _
                                   ByRef RawData() As Byte) As Byte()
  Dim PayloadSize As Long
  PayloadSize = CheckArray(RawData)
  If PayloadSize = 0 Then Exit Function

  Dim Out() As Byte
  ReDim Out(0 To PayloadSize + 11)

  ' Magic "SQUASH"
  Out(0) = &H53: Out(1) = &H51: Out(2) = &H55
  Out(3) = &H41: Out(4) = &H53: Out(5) = &H48
  Out(6) = CByte(CLng(Engine) ^ (1 / 8))   ' encode: 0->0, 256->2; decode: Data(6)^8
  Out(7) = CByte(Algorithm)                ' 12, 13, or 14
  CopyMemory VarPtr(Out(8)), VarPtr(UncompressedSize), 4
  CopyMemory VarPtr(Out(12)), VarPtr(RawData(0)), PayloadSize

  BuildSquashHeader = Out
End Function



