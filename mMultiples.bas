Attribute VB_Name = "mMultiples"
Option Explicit
Dim CurrentFirstRange As Integer, Trim As Integer, opTrim As Integer, opQTY As Integer, Cuts() As Integer, Ytrim As Boolean, Ymax As Boolean, opRange As Integer, currentdiff As Integer, opDiffRange As Integer
Dim Ythresh As Boolean, Thresh As Boolean, Max As Boolean, done As Boolean, Pos As Integer, opPos As Integer, opDiff() As Integer, Yqty As Boolean, Size As Integer, DiffQTY() As Integer, opDQTY As Integer, YdQTY As Boolean
Dim mult() As Variant, diff() As Integer, PrefTrim As Integer, Threshold() As Integer, MaxThreshold() As Integer, spec() As Integer, Qty As Integer, opSpecCombination() As Variant
Dim LastOpRange As Integer, DQTY As Integer, Ydiff As Boolean, Low As Boolean

Function FindOptimalSpecCombination(inputmult, InputDiff, InputPrefTrim, InputThresh, InputMaxThresh, InputSpec, InputDiffQTY, LowQTY)

Dim a As Integer, b As Integer, c As Integer

Ydiff = False

mult = inputmult
diff = InputDiff
PrefTrim = InputPrefTrim
Threshold = InputThresh
MaxThreshold = InputMaxThresh
spec = InputSpec
DiffQTY = InputDiffQTY

If LowQTY = True Then Low = True

ReDim opDiff(1 To UBound(diff))

For a = LBound(diff) To UBound(diff)
  currentdiff = a
  EmptyOP
  Yqty = False
   For b = LBound(mult) To UBound(mult)
     Do Until mult(b) = "Cuts" Or b = UBound(mult)
        b = b + 1
     Loop
             
        If b >= UBound(mult) Then Exit For
        CurrentFirstRange = b
        prepareinc
        If done = False Then proc
   Next b
  opDiff(currentdiff) = opRange
Next a

EmptyOP
YdQTY = False
Yqty = False
Ydiff = True
For a = LBound(opDiff) To UBound(opDiff)
done = False
If opDiff(a) = 0 Then done = True
If done = False Then
currentdiff = a
   For b = LBound(mult) To UBound(mult)
       If b = opDiff(a) Then Exit For
   Next b
   CurrentFirstRange = b
   prepareinc
   If done = False Then ProcA0
End If
Next a



a = opRange
Do Until mult(a) = "end"
  a = a + 1
Loop
        LastOpRange = a
        
Size = (LastOpRange + -opRange) + 5
ReDim opSpecCombination(1 To Size)
        
b = 1
For a = opRange To LastOpRange
    opSpecCombination(b) = mult(a)
    b = b + 1
Next a
  
   opSpecCombination(b) = "Difference"
   b = b + 1
   opSpecCombination(b) = diff(opDiffRange)
   b = b + 1
   opSpecCombination(b) = "DiffQTY"
   b = b + 1
   opSpecCombination(b) = opDQTY
   
   
                                FindOptimalSpecCombination = opSpecCombination
                
End Function

Sub proc()

If Ymax = False Then ProcA1

If Ymax = True Then
   If Max = True Then ProcA1

   If Max = False Then
     If opTrim >= PrefTrim * 2 And Trim < PrefTrim _
     Or Qty + 2 >= opQTY Then ProcA1
   End If
End If

End Sub

Sub ProcA0()

If YdQTY = False Then Store

If DQTY = opDQTY Then
   If Trim < opTrim And Qty <= opQTY Then Store
   If Trim < 10 And opTrim < 10 And Qty <= opQTY Then Store
   If opTrim > PrefTrim And Trim < PrefTrim Then Store
End If

If DQTY > opDQTY Then
  If Low = False Then
   If Trim <= opTrim And Qty < opQTY + 2 Then Store
   If Trim < 10 And opTrim < 10 And Qty < opQTY + 2 Then Store
  End If
  If Low = True Then Store
End If

If DQTY < opDQTY Then
  If Low = False Then
   If Trim < PrefTrim And opTrim > PrefTrim And Qty < opQTY + 2 Then Store
  End If
End If


End Sub

Sub ProcA1()
      
If Ytrim = False Then procA2

If Ytrim = True Then
   If Trim <= PrefTrim Then procA2
   
   If Trim > PrefTrim Then
      If Trim <= PrefTrim * 2 And Qty + 2 <= opQTY _
      And Max = True And Thresh = True _
      Or Trim <= PrefTrim * 2 And Qty + 2 <= opQTY _
      And Ymax = False And Ythresh = False Then procA2
   End If
End If
         
End Sub

Sub procA2()

If Ythresh = False Then
   If Ydiff = False Then procA3
   If Ydiff = True Then ProcA0
End If

If Ythresh = True Then
   If Thresh = True Then
      If Ydiff = False Then procA3
      If Ydiff = True Then ProcA0
   End If
   
   If Thresh = False Then
      If Qty < opQTY _
      Or opTrim <= PrefTrim * 2 _
      And Trim < PrefTrim Then
         If Ydiff = False Then procA3
         If Ydiff = True Then ProcA0
      End If
    End If
End If

End Sub

Sub procA3()

If Yqty = False Then Store

If Qty < opQTY Then
   If opTrim > 10 Then
       If Trim < opTrim * 1.5 And Trim < PrefTrim Then Store
   End If
   If Qty + 1 < opQTY And Trim < PrefTrim Then Store
   If Trim < 10 And opTrim < 10 Then Store
End If

If Qty = opQTY Then
   If Trim < opTrim Then Store
   If Trim = opTrim Then
      If Pos < opPos Then Store
   End If
End If

If Qty > opQTY Then
   If Qty < opQTY + 3 Then
      If opTrim >= PrefTrim * 2 And Trim < PrefTrim Then Store
      If Trim < 10 And opTrim >= 10 And Qty < opQTY + 2 Then Store
   End If
End If


End Sub


Sub Store()
Dim a As Integer, b As Integer, c As Integer

opDQTY = DQTY
opTrim = Trim
opQTY = Qty
opRange = CurrentFirstRange
opPos = Pos
opDiffRange = currentdiff
If Thresh = True Then Ythresh = True
If Thresh = False Then Ythresh = False
If Max = True Then Ymax = True
If Max = False Then Ymax = False
If Trim <= PrefTrim Then Ytrim = True
If Trim > PrefTrim Then Ytrim = False
Yqty = True
YdQTY = True

End Sub

Sub EmptyOP()

opDQTY = 0
opTrim = 0
opQTY = 0
opRange = 0
opPos = 0
Ythresh = False
Ymax = False
Ytrim = False

End Sub

Sub prepareinc()
Dim a As Integer, b As Integer, c As Integer, LastRange As Integer, Inputinc() As Integer, NumberOfCuts As Integer, d As Integer, Pmult() As Integer

DQTY = DiffQTY(currentdiff)
a = CurrentFirstRange
 Do Until mult(a) = "end"
   If mult(a) = "Cuts" Then
     Do Until mult(a + 1) = "Sum"
        a = a + 1
        b = b + 1
        ReDim Preserve Pmult(1 To b)
        Pmult(b) = mult(a)
     Loop
   End If
   a = a + 1
   If mult(a) = "Number Of Cuts" Then NumberOfCuts = mult(a + 1)
   If a >= UBound(mult) Then Exit Do
 Loop
 
                       LastRange = a

done = False
For a = CurrentFirstRange To LastRange
If a = LastRange Then Exit For
  If mult(a) = "Sum" Then
     If mult(a + 1) <= diff(currentdiff) Then
      Trim = diff(currentdiff) + -mult(a + 1)
      Else: done = True
    End If
  End If
Next a

If done = True Then Exit Sub

Cuts = SortInc(Pmult, NumberOfCuts)
Qty = NumberOfCuts

b = 1
Pos = 0
For a = LBound(Cuts) To UBound(Cuts)
    For b = LBound(spec) To UBound(spec)
        If Cuts(a) = spec(b) Then
          Pos = Pos + (((b * 0.1) * Cuts(a + 1)) + b)
        End If
    Next b
a = a + 1
Next a

Thresh = True
Max = True
For a = LBound(Cuts) To UBound(Cuts)
  For b = LBound(Threshold) To UBound(Threshold)
     If Cuts(a) = spec(b) Then
        c = -Cuts(a + 1) + Threshold(b)
        If c < 0 Then Thresh = False
        c = -Cuts(a + 1) + MaxThreshold(b)
        If c < 0 Then Max = False
    End If
  Next b
a = a + 1
Next a


      

End Sub

Function SortInc(Pmult, NumberOfCuts)
Dim a As Integer, b As Integer, c As Integer, Cuts() As Integer, cutQTY As Integer, good As Boolean, inc() As Integer

ReDim inc(1 To NumberOfCuts)

For a = 1 To UBound(Pmult)
If Pmult(a) = "Cuts" Then a = a + 1
If Pmult(a) = "Sum" Then
   Do Until mult(a) = "end"
      a = a + 1
   Loop
      If a + 2 < UBound(mult) Then
         a = a + 2
         Else: Exit For
      End If
End If
good = True
   For b = 1 To UBound(inc)
      If Pmult(a) = inc(b) Then good = False
   Next b
If good = True Then
c = c + 1
ReDim Preserve inc(1 To c)
inc(c) = Pmult(a)
End If
Next a

For a = LBound(inc) To UBound(inc)
   If a + 1 > UBound(inc) Then Exit For
   If inc(a + 1) = 0 Then Exit For
Next a
        ReDim Preserve inc(1 To a)
        ReDim Cuts(1 To a * 2)

c = 0
For a = LBound(inc) To UBound(inc)
cutQTY = 0
   For b = LBound(Pmult) To UBound(Pmult)
       If Pmult(b) = inc(a) Then cutQTY = cutQTY + 1
   Next b
c = c + 1
Cuts(c) = inc(a)
c = c + 1
Cuts(c) = cutQTY
Next a


                    SortInc = Cuts

End Function


