Attribute VB_Name = "modVBASquash_Test"
Option Explicit

' TESTS WRITTEN BY GOOGLE GEMINI 2.5 PRO PREVIEW

Sub RunFullCompressionTestSuite()
    Dim startTime As Double
    startTime = Timer

    Debug.Print "======================================================"
    Debug.Print "STARTING vbaSquash COMPRESSION TEST SUITE"
    Debug.Print "Timestamp: " & Now
    Debug.Print "======================================================"

    TestSpecificAlgorithm MSZIP, "MSZIP"
    TestSpecificAlgorithm XPRESS, "XPRESS"
    TestSpecificAlgorithm XPRESS_HUFF, "XPRESS_HUFF"
    TestSpecificAlgorithm LZMS, "LZMS"

    TestEdgeCases

    Debug.Print "======================================================"
    Debug.Print "vbaSquash COMPRESSION TEST SUITE COMPLETE"
    Debug.Print "Total time: " & Format(Timer - startTime, "0.000") & " seconds"
    Debug.Print "======================================================"
End Sub

Private Sub TestSpecificAlgorithm(ByVal currentAlgorithm As COMPRESS_ALGORITHM_ENUM, ByVal algoName As String)
    Dim testStrings As Variant
    Dim testData() As Byte
    Dim CompressedData() As Byte
    Dim DecompressedData() As Byte
    Dim i As Long
    Dim s As String
    Dim originalSize As Long, compSize As Long, decompSize As Long
    Dim overallSuccess As Boolean
    Dim testCaseSuccess As Boolean

    Debug.Print vbCrLf & "--- Testing Algorithm: " & algoName & " (" & currentAlgorithm & ") ---"

    ' Diverse set of strings to test different data patterns and sizes
    testStrings = Array( _
        String$(1, "A"), _
        "Hello", _
        "The quick brown fox jumps over the lazy dog.", _
        String$(10, "Repeat "), _
        String$(50, "A"), _
        String$(100, Chr$(0) & Chr$(255) & "Mid"), _
        String$(250, "TestPattern123"), _
        String$(512, "BufferMode Test String"), _
        String$(1024, "A"), _
        String$(4096, "B"), _
        String$(8192, "C"), _
        String$(16384, "D"), _
        String$(32768 - 10, "E"), _
        String$(32768 + 10, "F"), _
        String$(65536, "G") _
    )

    overallSuccess = True

    For i = LBound(testStrings) To UBound(testStrings)
        s = CStr(testStrings(i))
        testData = StrConv(s, vbFromUnicode) ' Convert to byte array (system ANSI)
        originalSize = GetArrayElementCount(testData)
        testCaseSuccess = True

        Debug.Print "  Test Case " & i + 1 & ": Input Length (Chars) = " & Len(s) & ", Original Byte Size = " & originalSize

        ' --- Compression ---
        Erase CompressedData
        On Error Resume Next ' Catch any direct errors from CompressBytes
        CompressedData = CompressBytes(testData, currentAlgorithm)
        If Err.Number <> 0 Then
            Debug.Print "    COMPRESSION: FAILED (Runtime Error " & Err.Number & " in CompressBytes: " & Err.Description & ")"
            compSize = -1
            testCaseSuccess = False
            Err.Clear
        Else
            If IsArrayPopulated(CompressedData) Then
                compSize = GetArrayElementCount(CompressedData)
                Debug.Print "    COMPRESSION: Success. Compressed Size = " & compSize & " (" & Format(100 * compSize / originalSize, "0.0") & "%)"
                If compSize > 0 And compSize >= 12 Then ' Check if large enough for header
                    If CompressedData(LBound(CompressedData)) = &HA And CompressedData(LBound(CompressedData) + 1) = &H51 Then
                        Debug.Print "      Header Signature (0A 51..): Present. Algo in header: " & CompressedData(LBound(CompressedData) + 7)
                    Else
                        Debug.Print "      Header Signature (0A 51..): NOT PRESENT or malformed!"
                    End If
                ElseIf compSize > 0 Then
                     Debug.Print "      Output too short for full header."
                End If
            Else
                compSize = 0
                Debug.Print "    COMPRESSION: FAILED (CompressBytes returned empty/uninitialized array)."
                testCaseSuccess = False
            End If
        End If
        On Error GoTo 0

        ' --- Decompression ---
        If testCaseSuccess And compSize > 0 Then
            Erase DecompressedData
            Dim detectedAlgo As COMPRESS_ALGORITHM_ENUM
            detectedAlgo = IsCompressed(CompressedData) ' Test IsCompressed

            If detectedAlgo <> currentAlgorithm Then
                Debug.Print "    DECOMPRESSION: IsCompressed detected " & detectedAlgo & ", expected " & currentAlgorithm & ". Proceeding with original algo."
                ' For this test, we force decompression with the known original algorithm.
                ' In real use, you might trust IsCompressed or handle mismatches.
            End If
            
            On Error Resume Next
            ' Using the original algorithm for decompression in this test for directness
            DecompressedData = DecompressBytes(CompressedData, currentAlgorithm)
            If Err.Number <> 0 Then
                Debug.Print "    DECOMPRESSION: FAILED (Runtime Error " & Err.Number & " in DecompressBytes: " & Err.Description & ")"
                decompSize = -1
                testCaseSuccess = False
                Err.Clear
            Else
                If IsArrayPopulated(DecompressedData) Then
                    decompSize = GetArrayElementCount(DecompressedData)
                    Debug.Print "    DECOMPRESSION: Success. Decompressed Size = " & decompSize
                    If decompSize = originalSize Then
                        If VerifyByteArrays(testData, DecompressedData) Then
                            Debug.Print "      Data Verification: MATCHES ORIGINAL"
                        Else
                            Debug.Print "      Data Verification: !!! MISMATCH CONTENT !!!"
                            testCaseSuccess = False
                        End If
                    Else
                        Debug.Print "      Data Verification: !!! MISMATCH SIZE (Original: " & originalSize & ", Decompressed: " & decompSize & ") !!!"
                        testCaseSuccess = False
                    End If
                Else
                    decompSize = 0
                    Debug.Print "    DECOMPRESSION: FAILED (DecompressBytes returned empty/uninitialized array)."
                    testCaseSuccess = False
                End If
            End If
            On Error GoTo 0
        ElseIf testCaseSuccess And compSize = 0 Then
             Debug.Print "    DECOMPRESSION: Skipped (Compression resulted in 0 bytes)."
        Else
            Debug.Print "    DECOMPRESSION: Skipped (Compression failed)."
        End If
        
        If Not testCaseSuccess Then overallSuccess = False
        Debug.Print "  -------------------------------------------"
    Next i
    Debug.Print "--- Algorithm " & algoName & " Test Summary: " & IIf(overallSuccess, "ALL PASSED", "ONE OR MORE FAILED") & " ---"
End Sub

Private Sub TestEdgeCases()
    Dim emptyArr() As Byte
    Dim tinyArr(0 To 0) As Byte
    Dim CompressedData() As Byte
    Dim DecompressedData() As Byte

    Debug.Print vbCrLf & "--- Testing Edge Cases ---"

    ' --- Test 1: Compress Empty Array ---
    Debug.Print "  Edge Case 1: Compressing Empty Array (MSZIP)"
    Erase CompressedData
    CompressedData = CompressBytes(emptyArr, MSZIP)
    If IsArrayPopulated(CompressedData) Then
        Debug.Print "    Compress Empty: FAILED (Expected empty, Got " & GetArrayElementCount(CompressedData) & " bytes)"
    Else
        Debug.Print "    Compress Empty: PASSED (Correctly returned empty/uninitialized)"
    End If

    ' --- Test 2: Decompress Empty Array ---
    Debug.Print "  Edge Case 2: Decompressing Empty Array (MSZIP)"
    Erase DecompressedData
    DecompressedData = DecompressBytes(emptyArr, MSZIP)
    If IsArrayPopulated(DecompressedData) Then
        Debug.Print "    Decompress Empty: FAILED (Expected empty, Got " & GetArrayElementCount(DecompressedData) & " bytes)"
    Else
        Debug.Print "    Decompress Empty: PASSED (Correctly returned empty/uninitialized)"
    End If

    ' --- Test 3: Decompress Invalid Data (not matching header) ---
    Debug.Print "  Edge Case 3: Decompressing Invalid Data (MSZIP)"
    tinyArr(0) = &H12 ' Just some random byte
    Erase DecompressedData
    DecompressedData = DecompressBytes(tinyArr, MSZIP)
    If IsArrayPopulated(DecompressedData) Then
        Debug.Print "    Decompress Invalid: FAILED (Expected empty, Got " & GetArrayElementCount(DecompressedData) & " bytes)"
    Else
        Debug.Print "    Decompress Invalid: PASSED (Correctly returned empty/uninitialized or handled error)"
    End If
    
    ' --- Test 4: IsCompressed on Empty/Invalid Data ---
    Dim algoCheck As COMPRESS_ALGORITHM_ENUM
    Debug.Print "  Edge Case 4: IsCompressed on Empty Array"
    algoCheck = IsCompressed(emptyArr)
    If algoCheck = 0 Then ' Assuming 0 means not compressed / invalid
        Debug.Print "    IsCompressed Empty: PASSED (Returned " & algoCheck & ")"
    Else
        Debug.Print "    IsCompressed Empty: FAILED (Returned " & algoCheck & ")"
    End If

    Debug.Print "  Edge Case 5: IsCompressed on Invalid Data (1 byte)"
    algoCheck = IsCompressed(tinyArr)
     If algoCheck = 0 Then
        Debug.Print "    IsCompressed Invalid (1 byte): PASSED (Returned " & algoCheck & ")"
    Else
        Debug.Print "    IsCompressed Invalid (1 byte): FAILED (Returned " & algoCheck & ")"
    End If
    Debug.Print "--- Edge Cases Test Complete ---"
End Sub


' --- Helper Functions for Test Module ---
Private Function IsArrayPopulated(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayPopulated = False
    If IsArray(arr) Then
        Dim lb As Long, ub As Long
        lb = LBound(arr)
        ub = UBound(arr)
        If Err.Number = 0 Then
            If lb <= ub Then IsArrayPopulated = True
        End If
        Err.Clear
    End If
End Function

Private Function GetArrayElementCount(arr As Variant) As Long
    ' Returns number of elements if populated, 0 otherwise
    On Error Resume Next
    GetArrayElementCount = 0
    If IsArrayPopulated(arr) Then
        GetArrayElementCount = UBound(arr) - LBound(arr) + 1
    End If
    Err.Clear
End Function

Private Function VerifyByteArrays(arr1() As Byte, arr2() As Byte) As Boolean
    Dim i As Long
    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long

    If Not IsArrayPopulated(arr1) And Not IsArrayPopulated(arr2) Then
        VerifyByteArrays = True ' Both empty, considered match
        Exit Function
    End If
    If Not IsArrayPopulated(arr1) Or Not IsArrayPopulated(arr2) Then
        VerifyByteArrays = False ' One empty, one not, mismatch
        Exit Function
    End If

    lb1 = LBound(arr1): ub1 = UBound(arr1)
    lb2 = LBound(arr2): ub2 = UBound(arr2)

    If (ub1 - lb1) <> (ub2 - lb2) Then
        VerifyByteArrays = False ' Different number of elements
        Exit Function
    End If

    VerifyByteArrays = True ' Assume match until proven otherwise
    For i = lb1 To ub1
        If arr1(i) <> arr2(i - lb1 + lb2) Then ' Compare elements, adjusting for potentially different LBounds
            VerifyByteArrays = False
            Exit Function
        End If
    Next i
End Function

Private Function BytesToHex(bytes As Variant, Optional MaxBytesToShow As Long = 16) As String
    Dim k As Long, s As String, ub As Long, lb As Long, count As Long
    Dim bArr() As Byte

    If Not IsArrayPopulated(bytes) Then
        BytesToHex = "[Not Populated/Empty]"
        Exit Function
    End If
    
    bArr = bytes

    lb = LBound(bArr)
    ub = UBound(bArr)
    
    count = ub - lb + 1
    If MaxBytesToShow > -1 And MaxBytesToShow < count Then ' Allow MaxBytesToShow = -1 for all
        count = MaxBytesToShow
    End If
    If MaxBytesToShow = 0 Then count = 0 ' Show 0 bytes if explicitly asked

    If count <= 0 Then BytesToHex = "[Zero Length or Empty]": Exit Function

    For k = lb To lb + count - 1
        If k > UBound(bArr) Then Exit For ' Safety break if count was > actual elements
        s = s & Right("0" & Hex(bArr(k)), 2) & " "
    Next k
    BytesToHex = Trim(s)
End Function

