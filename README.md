# vbaSquash
## Compression routines for VBA

`vbaSquash` leverages the built-in `cabinet.dll` functions available on Windows 8+, providing access to compression algorithms MSZIP, XPRESS, XPRESS_HUFF, and LZMS without external dependencies.

### 🛠️ Features:

*   **Simple Interface:** Easy-to-use functions like `CompressBytes`, `DecompressBytes`, `CompressFile`, and `DecompressFile`.
*   **No additional software:** Uses Windows' own compression routines.
*   **Automatic Algorithm Detection (for files compressed by vbaSquash):** The library attempts to auto-detect the algorithm used if you're decompressing a file it previously compressed. This differs from other methods.


### 🚀 Basic Usage

#### Compress a File
```vba
Dim success As Boolean
success = CompressFile("C:\input.txt", , LZMS)
' Output will be saved as C:\input.txt.compressed
```

#### Decompress a File
```vba
Dim success As Boolean
success = DecompressFile("C:\input.txt.compressed", "C:\NewFile.txt")
' Output will be saved as C:\NewFile.txt
```

#### Compress a byte array
```vba
Dim inputData() As Byte
inputData = ReadFile("C:\file.dat")

Dim compressedData() As Byte
compressedData = CompressBytes(inputData)
```

### 📝 Side-Note:

This was a rabbit hole.

When you use `vbaSquash` (or the underlying Windows Compression API in its default "Buffer Mode"), a small, neat 12-byte header appears to be automatically added to your compressed data. This header is handy: it includes a "magic number" (`0A 51 E5 C0 18 00`), an ID for the compression algorithm used (#2 to #5), and the original size of your uncompressed data.

This all came about while trying to finalise another project re: file format inference/detection and the compressed files from these vbaSquash routines were added to the test dataset. I noticed the repeating pattern (the magic number), and then I noticed that the 8th byte always seemed to be a number between 2 and 5, which 'coincidentally' seemed to match the algorithm enumeration in cabinet.dll.

I asked Google Gemini about this 8th byte because I couldn't find any information on it. Google Gemini said I was wrong. I compressed 100 more files, and they all came back with the same results - I updated Google Gemini with the results. Google Gemini was unmoved. Back before Google neutered Gemini by making its internal monologue bland and utterly useless, I saw that Gemini reasoned that the results must be because the user (i.e., me) is accidentally doing something to the compression code that leads to these results. So I asked Gemini to come up with tests - the results were in, and I was right. Yay Team Human.

### Choosing an Algorithm:

*   **XPRESS:** Generally very fast with a medium compression ratio. Good all-rounder. This is my default.
*   **MSZIP:** Good compression. Well-known due to CAB files. Apparently quite legacy, and from personal experience, it's a bit 'odd'.
*   **LZMS:** Can offer high compression ratios, especially for larger data (>2MB), but is slower to compress.
*   **XPRESS_HUFF:** A variation of XPRESS, often similar performance.

I do not take a view on speed, as much of what I compress/decompress is not especially large such that I would notice the difference.

### License:

MIT
