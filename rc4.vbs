'
'  The MIT License:
'
'  Copyright (c) June 2012 Kevin Devine
'
'  Permission is hereby granted,  free of charge,  to any person obtaining a copy
'  of this software and associated documentation files (the "Software"),  to deal
'  in the Software without restriction,  including without limitation the rights
'  to use,  copy,  modify,  merge,  publish,  distribute,  sublicense,  and/or sell
'  copies of the Software,  and to permit persons to whom the Software is
'  furnished to do so,  subject to the following conditions:
'
'  The above copyright notice and this permission notice shall be included in
'  all copies or substantial portions of the Software.
'
'  THE SOFTWARE IS PROVIDED "AS IS",  WITHOUT WARRANTY OF ANY KIND,  EXPRESS OR
'  IMPLIED,  INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,  DAMAGES OR OTHER
'  LIABILITY,  WHETHER IN AN ACTION OF CONTRACT,  TORT OR OTHERWISE,  ARISING FROM,
'  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
'  THE SOFTWARE.
'
    Const DEBUG_FLAG = False

    Dim rc4_ctx, ciphertext, plaintext, key, decrypted
    key = Array(&H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF)
    plaintext = "Hello, World!"
    
    Set rc4_ctx = New RC4

    ' initialize key and dump
    Call rc4_ctx.SetKey(key, UBound(key) + 1)
    Call DumpData("Key", rc4_ctx.Data)

    ' encrypt plaintext and dump
    Call DumpData("Plaintext", plaintext)
    Call rc4_ctx.Encrypt(plaintext, Len(plaintext), ciphertext)
    Call DumpData("Ciphertext", ciphertext)
    
    ' initialize key again but don't dump
    Call rc4_ctx.SetKey(key, UBound(key) + 1)
    
    ' decrypt ciphertext and dump
    Call rc4_ctx.Decrypt(ciphertext, UBound(ciphertext), decrypted)    
    Call DumpData("Decrypted", decrypted)

    WScript.Echo
    WScript.Quit
  
    ' ***********************************************************
    '
    ' display the contents of data array
    ' mainly for debugging purposes
    '
    ' ***********************************************************
    Public Sub DumpData(ByVal txt, ByVal data)
      
      WScript.Echo vbCrLf & vbCrLf & txt & ":"
      Dim data_len
      
      If (VarType(data) = vbString) Then
        data_len = Len(data)
      Else
        data_len = UBound(data)
      End If
      
      Dim i
      
      For i = 0 To data_len - 1
        If (i Mod 16 = 0) Then
          WScript.Echo
        End If
        Dim p
        
        If (VarType(data) = vbString) Then
          p = Asc( Mid(data, (i + 1), 1))
          p = Hex(p)
        Else
          p = Hex(data(i))
        End If
        
        If (Len(p) < 2) Then
          p = "0" & p
        End If
        WScript.StdOut.Write p & " "
      Next

    End Sub
  
'  
'
' RC4 Encryption
'
'
Class RC4
    Public x, y
    Public data
    
    Private Sub Class_Initialize
      ReDim data(256)
    End Sub
    
    ' ***********************************************************
    '
    ' initialise RC4 key
    '
    ' ***********************************************************
    Public Sub SetKey(ByVal keyData, ByVal keyLen)
      Dim i
      
      x = 0
      y = 0
      
      If (DEBUG_FLAG) Then
        WScript.StdOut.WriteLine(vbCrLf & "Key Size is " & keyLen)
      End If
      
      For i = 0 To UBound(data) - 1
        data(i) = i
      Next
      
      Dim tmp, id1, id2
      id1 = 0
      id2 = 0
      
      For i = 0 To UBound(data) - 1
        tmp = data(i)
        
        Dim c
        
        If (VarType(keyData) = (vbArray + vbVariant)) Then
          c = keyData(id1)
        Else
          c = CInt(Asc(Mid(keyData, (id1 + 1), 1)))
        End If
        
        id2 = Add8(c, Add8(tmp,id2))
        id1 = id1 + 1
        
        If (id1 = keyLen) Then
          id1 = 0
        End If
        
        data(i) = data(id2)
        data(id2) = tmp
      Next
    End Sub
        
    ' ***********************************************************
    '
    ' encrypt binary data
    '
    ' ***********************************************************
    Public Sub Encrypt(ByVal input, ByVal inputLen, ByRef output)
      Dim i, tx, ty, c
      ReDim output(inputLen)
      
      If (DEBUG_FLAG) Then
        WScript.StdOut.WriteLine(vbCrLf & "Size of plaintext is " & inputLen)
      End If
      
      For i = 0 To inputLen - 1
        x = Add8(x, 1)
        tx = data(x)
        
        y = Add8(tx, y)
        ty = data(y)
        
        data(x) = ty
        data(y) = tx
        
        If (VarType(input) = vbString) Then
          c = CInt(Asc(Mid(input, (i + 1), 1)))
        Else
          c = input(i)
        End If
        
        output(i) = data(Add8(tx, ty)) Xor c            
      Next
    End Sub
    
    ' ***********************************************************
    '
    ' encrypt binary data
    '
    ' ***********************************************************
    Public Sub Decrypt(ByVal input, ByVal inputLen, ByRef output)
      Call Encrypt(input, inputLen, output)
    End Sub
    
    ' ***********************************************************
    '
    ' add 2 8-bit unsigned values
    '
    ' ***********************************************************
    Private Function Add8(ByRef lX, ByRef lY)
        Dim lX4, lY4, lX8, lY8
        Dim lResult
     
        lX = CInt(lX)
        lY = CInt(lY)
        
        lX8 = lX And &H80
        lY8 = lY And &H80
        lX4 = lX And &H40
        lY4 = lY And &H40
     
        lResult = (lX And &H3F) + (lY And &H3F)
     
        If (lX4 And lY4) Then
          lResult = lResult Xor &H80 Xor lX8 Xor lY8
        ElseIf (lX4 Or lY4) Then
          If (lResult And &H40) Then
            lResult = lResult Xor &HC0 Xor lX8 Xor lY8
          Else
            lResult = lResult Xor &H40 Xor lX8 Xor lY8
          End If
        Else
          lResult = lResult Xor lX8 Xor lY8
        End If
        
        Add8 = CInt(lResult)
    End Function

End Class
