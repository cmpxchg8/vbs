'
'  The MIT License:
'
'  Copyright (c) 2012 Kevin Devine
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


' *********************************************************
' C:\>cscript /nologo bfe.vbs 11 11 16 8 15000000
'
' Number of CPU = 8
' Total keys    = 17,592,186,044,416
' Average Speed = 15,000,000 per second.
'
' ETA: 1 day(s) 16 hour(s) 43 minute(s) 22 second(s)
'
' *********************************************************

Set argv = WScript.Arguments
argc = WScript.Arguments.Count

If (argc <> 5) Then
  print vbCrLf
  print "BFE - Brute Force Estimator" & vbCrLf & vbCrLf
  print "Usage: bfe.vbs <min_len> <max_len> <alpha_len> <max_cpu> <avg_speed>" & vbCrLf
  quit
Else
  min_pwd   = argv(0)
  max_pwd   = argv(1)
  alpha_len = argv(2)
  cpu       = argv(3)
  avg_speed = argv(4)
  
  total = 0
  
  For i = min_pwd To max_pwd
    total = total + (alpha_len ^ i)
  Next
  
  sec = total / avg_speed / cpu
  mns = 0
  hrs = 0
  days = 0
  
  If (sec >= 60) Then
    mns = sec \ 60
	sec = sec Mod 60
  End If
  
  If (mns >= 60) Then
    hrs = mns \ 60
	mns = mns Mod 60
  End If
  
  If (hrs >= 24) Then
    days = hrs \ 24
	hrs = hrs Mod 24
  End If
  
  print vbCrLf
  print "Number of CPU = " & cpu & vbCrLf
  print "Total keys    = " & FormatNumber(total,0,0,-1) & vbCrLf
  print "Average Speed = " & FormatNumber(avg_speed,0,0,-1) & " per second." & vbCrLf & vbCrLf
  
  print "ETA: " & days & " day(s) "
  print hrs & " hour(s) " & mns  & " minute(s) "
  print sec & " second(s)" & vbCrLf
  
  quit
End If

Sub print(s)
  WScript.StdOut.Write s
End Sub

Sub quit()
  WScript.Quit
End Sub
