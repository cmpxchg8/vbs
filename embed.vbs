'
'  The MIT License:
'
'  Copyright (c) 2011, 2012 Kevin Devine
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
  Option Explicit

  Const ERROR_SUCCESS = 0
  
  Const ForReading    = 1
  Const ForWriting    = 2

  Call main(WScript.Arguments.Count, WScript.Arguments)
  WScript.Quit()
  
  Sub main(ByVal argc, ByRef argv)    
    
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    
    If (argc >= 1) Then
      Dim vbs_file : vbs_file = argv(0)
      ' ensure VBS file exists
      If (fso.FileExists(vbs_file)) Then
        ' ensure it has VBS extension
        Dim ext : ext = LCase(fso.GetExtensionName(vbs_file))
        If (ext = "vbs") Then
          printf(Array("\n[+] Processing %s", vbs_file))
          
          ' try open the file
          Dim vbs_in : Set vbs_in = fso.OpenTextFile(vbs_file, ForReading)
          
          If (Err = ERROR_SUCCESS) Then
            ' get name without extension
            Dim base_name : base_name = fso.GetBaseName(vbs_file)
            Dim temp_name : temp_name = base_name & "_TEMP.VBS"
            
            ' create and overwrite if necessary
            Dim cmd_out : Set cmd_out = fso.CreateTextFile(base_name & ".cmd", True)
            
            If (Err = ERROR_SUCCESS) Then
              Call fprintf(cmd_out, Array("@Echo Off\n"))
              Call fprintf(cmd_out, Array("If Exist %s DEL %s\n", temp_name, temp_name))
              Call fprintf(cmd_out, Array(":' VBSCRIPT STARTS HERE - DO NOT EDIT!\n"))
              
              ' read VBS and save to CMD format, removing Option Explicit if found
              Do Until vbs_in.AtEndOfStream
                Dim vbs_line : vbs_line = vbs_in.ReadLine()
                
                If (StrComp(UCase(Trim(vbs_line)), UCase("Option Explicit"), 1) <> 0) Then
                  cmd_out.WriteLine(":" & vbs_line)
                End If
              Loop
              
              vbs_in.Close()
              
              Dim find_str
              
              find_str = "FINDSTR " & Chr(34) & "^:" & Chr(34) & " " & Chr(34) & "%~sf0" & Chr(34)
              find_str = find_str & " > " & temp_name & " & CSCRIPT //T:300 //nologo " & temp_name
              
              cmd_out.Write(find_str)
              
              ' finally whatever arguments were specified..
              If (argc = 2) Then
                Dim i, arg_count
                arg_count = argv(1)
                
                If (arg_count > 9) Then
                  printf(Array("\n[-] WARNING: Batch files are limited to 9 arguments without command extensions."))
                  arg_count = 9
                End If
                  
                For i = 0 To arg_count - 1
                  cmd_out.Write(" %" & (i + 1))
                Next
              End If
              
              cmd_out.Write(" & DEL " & temp_name)
              cmd_out.Close
              
              printf(Array("\n[+] Converted %s to %s successfully\n", vbs_file, base_name & ".cmd"))
            Else
              printf(Array("\n[-] Unable to create %s", base_name & ".cmd"))
            End If
          Else
            printf(Array("\n[-] Unable to open %s\n", vbs_file))
          End If
        Else
          printf(Array("\n[-] Missing .VBS extension : %s\n", vbs_file))
        End If
      Else
        printf(Array("\n[-] File does not exist : %s\n", vbs_file))
      End If
    Else
      printf(Array("\nEmbed v0.1 - Embed a VBS file in BATCH/CMD file"))
      printf(Array("\nUSAGE: Embed <vbs file> <number of parameters>\n"))
    End If
  End Sub
  
  '
  '
  '
  Function printf(ByRef args)
    Dim s
    Call sprintf(s, args)
    WScript.StdOut.Write(s)
  End Function
   
  '
  ' write printf strings to file
  '
  Function fprintf(ByVal fd, ByRef args)
    Dim s
    Call sprintf(s, args)
    
    fd.Write(s)
  End Function

  ' ##################################################################################
  '
  ' minimal printf emulation for VBScript
  ' WARNING: Not a full implementation or bug free..misuse and you'll easily break it.
  '
  ' supported specifiers
  '
  ' d or i   = decimal or integer
  ' c        = character
  ' s        = string
  ' x        = lowercase hex byte
  ' X        = uppercase hex byte
  ' o        = octal byte
  ' %        = percentage sign
  '
  ' supported flags
  '
  ' -        = left justify
  ' #        = precede with 0, 0x or 0X if radix
  ' 0        = pad with zeros
  '
  ' supported width
  '
  ' *        = specified in argument preceding value
  ' (number) = number of spaces or zeros to pad
  '
  ' supported escape sequences
  '
  ' n        = line feed
  ' t        = tab
  ' r        = carriage return
  ' \        = backslash
  '
  ' Copyright (c) 2012 - Kevin Devine
  '
  ' #################################################################################
  Function sprintf(ByRef s1, ByRef args)
    
    Const RIGHT_JUSTIFY = 0
    Const LEFT_JUSTIFY  = 1
     
    ' array and not empty?
    If (IsArray(args) And UBound(args) <> -1) Then
      Dim s2

      s1 = args(0)
       
      ' ensure atleast 1 parameter before parsing
      If (UBound(args) >= 0) Then
        Dim fmt, i, index, pad_len
        fmt = s1
        index = 1

        For i = 1 To Len(fmt)

          Dim justify, zero_pad, lower_case, arg, show_radix

          justify    = RIGHT_JUSTIFY    ' default
          zero_pad   = False
          lower_case = False
          show_radix = False
          pad_len    = 0
          arg        = ""
           
          ' ############################################# SPECIFIER?
          ' Check if beginning of specifier
          '
          Dim c : c = Mid(fmt, i, 1)
           
          If (c = "%") Then
          
            ' skip percentage sign
            i = i + 1
               
            ' ############################################# step 1
            ' Check for supported flags
            '
            c = Mid(fmt, i, 1)
               
            ' show radix?
            If (c = "#") Then
              show_radix = True
              i = i + 1
                 
            ' left justify?
            ElseIf (c = "-") Then
              justify = LEFT_JUSTIFY
              i = i + 1
               
            ' pad with zeros instead of spaces?             
            ElseIf (c = "0") Then
              zero_pad = True
              i = i + 1
              
            ' space?  
            ElseIf (c = " ") Then
              arg = arg & " "
            End If
              
            ' ############################################# step 2
            ' Check for padding width
            '
            c = Mid(fmt, i, 1)
            
            ' specified in arguments?             
            If (c = "*") Then
              pad_len = args(index)
              index = index + 1
              i = i + 1
            
            ' if it's a number presume it to be specified length
            ElseIf (IsNumeric(c)) Then
              Do While True
                c = Mid(fmt, i, 1)

                ' convert string to binary
                If (IsNumeric(c)) Then
                  pad_len = pad_len * 10
                  pad_len = pad_len + (c - CInt("0"))
                  i = i + 1
                Else
                  Exit Do
                End If
              Loop   
            End If
               
            ' ############################################# step 3
            ' What specifier?
            '
            c = Mid(fmt, i, 1)

            Select Case c
              Case "s"
                arg = arg & args(index)
              Case "c"
                arg = arg & Chr(args(index))
              Case "d"
                arg = arg & CStr(args(index))
              Case "i"
                arg = arg & CStr(args(index))
              Case "x"
                If (show_radix) Then
                  arg = arg & "0x"
                End If
                arg = arg & LCase(Hex(args(index)))
              Case "X"
                If (show_radix) Then
                  arg = "0X"
                End If
                  arg = arg & UCase(Hex(args(index)))
              Case "o"
                If (show_radix) Then
                  arg = "0"
                End If
                  arg = arg & Oct(args(index))
              Case "%"
                arg = "%"
              Case Else
                WScript.Echo "Unrecognized specifier: " & c
            End Select
            
            ' ############################################# step 4
            ' process padding length and justifcation
            '
            If (pad_len <> 0) Then
              If (pad_len > Len(arg)) Then
                pad_len = pad_len - Len(arg)
              Else
                pad_len = 0
              End If
            End If

            ' pad with zeros or spaces?
            Dim pad_str
            
            If (zero_pad) Then
              pad_str = String(pad_len, "0")
            Else
              pad_str = Space(pad_len)
            End If

            ' justify left or right?
            If (justify = LEFT_JUSTIFY) Then
              s2 = s2 & arg & pad_str
            Else
              s2 = s2 & pad_str & arg
            End If

            ' display percent?
            If (arg <> "%") Then
              index = index + 1
            End If
            
          ' ############################################# ESCAPE SEQUENCE?
          ' 
          '
          ElseIf (c = "\") Then
             
            ' get the byte
            Dim esc : esc = Mid(fmt, i + 1, 1)
             
            ' newline?
            If (esc = "n") Then
              s2 = s2 & vbCrLf
              i = i + 1
            
            ' carriage return?           
            ElseIf (esc = "r") Then
              s2 = s2 & vbCr
              i = i + 1
            
            ' horizontal tab?
            ElseIf (esc = "t") Then
              s2 = s2 & vbTab
              i = i + 1
            
            ' backslash?
            ElseIf (esc = "\") Then
              s2 = s2 & "\"
              i = i + 1
            
            Else
              ' unrecognized ..
              s2 = s2 & "\"
            End If
          ' ############################################# IGNORE
          '
          Else         
            s2 = s2 & c
          End If
        Next
        s1 = s2
      End If
    Else
      WScript.Echo "No Array provided"
    End If
  End Function
