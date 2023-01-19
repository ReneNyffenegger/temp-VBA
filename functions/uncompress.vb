Public Declare Function LZUncompress Lib "kernel32" Alias "LZUncompress" (lpSrc As Any,  lpDst As Any, dwSrcLen As Long) As Long

Public Function UncompressData(ByRef CompressedData() As Byte) As Byte()

    Dim UncompressedData() As Byte
    Dim UncompressedSize As Long
    UncompressedSize = LZUncompress(CompressedData(0), 0, UBound(CompressedData) + 1)
    If UncompressedSize > 0 Then
        ReDim UncompressedData(UncompressedSize - 1)
        LZUncompress CompressedData(0), UncompressedData(0), UBound(CompressedData) + 1
    Else
        ReDim UncompressedData(0)
    End If
    UncompressData = UncompressedData

End Function

