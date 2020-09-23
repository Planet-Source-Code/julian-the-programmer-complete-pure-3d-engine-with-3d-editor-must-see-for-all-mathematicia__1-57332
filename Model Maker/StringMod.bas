Attribute VB_Name = "StringMod"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Const LPTR = (&H0 Or &H40)

Public Function MemAlloc(Strin As String) As Long
    Dim PointerA As Long, lSize As Long
    
    lSize = LenB(Strin) 'Length of string in bytes.
    
    'Allocate the memory needed and returns a pointer to that memory
    PointerA = LocalAlloc(LPTR, lSize + 4)
    If PointerA <> 0 Then
        'Final allocation
        CopyMemory ByVal PointerA, lSize, 4
        If lSize > 0 Then
            'copy the string to that allocated memory.
            CopyMemory ByVal PointerA + 4, ByVal StrPtr(Strin), lSize
        End If
    End If
    'return the pointer to the string stored memory
    MemAlloc = PointerA
End Function

Public Function RetMemory(PointerA As Long) As String

    Dim lSize As Long, sThis As String
    If PointerA = 0 Then
        GetMemory = ""
    Else
        'get the size of the string stored at pointer "PointerA"
        CopyMemory lSize, ByVal PointerA, 4
       If lSize > 0 Then
            'buffer a varible
            sThis = String(lSize \ 2, 0)
            'retrive the data at the address of "PointerA"
            CopyMemory ByVal StrPtr(sThis), ByVal PointerA + 4, lSize
            'return the buffer
            RetMemory = sThis
        End If
    End If
End Function

Public Sub FreeMemory(PointerA As Long)
    'frees up the memory at the address of PointerA
    LocalFree PointerA
End Sub

Public Function OpenFile(fInStream As String) As String
  On Error GoTo err
  Dim I As Long, strText As String
  I = FreeFile
  strText = ""
  Open fInStream For Input Lock Write As #I
  DoEvents
  strText = StrConv(InputB$(LOF(I), I), vbUnicode)
  Close #I
  OpenFile = strText
  Exit Function
err:
  MsgBox "An error accourd while trying to load the following file: '" & fInStream & "'", vbCritical
End Function

