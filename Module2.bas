Attribute VB_Name = "Module2"
Option Explicit


Function FindAllCombinations(InputArr, InputLD, InputLowD)

Dim a As Integer, b As Integer, c As Integer, TotalCombinations As Long, i As Long, y As Integer, k As Integer, e As Integer, F As Long, IncDiff As Long, r As Integer, cNumber As Long, Hcut As Integer
Dim h As Integer, ii As Integer, BlockArr() As Integer, g As Integer, First As Boolean, Lnumber As Boolean, p As Integer, sum As Long, CurrentLastLocation As Long, NumberOfCuts As Integer
Dim calc As Boolean, one As Integer, one2 As Integer, TotalCuts As Integer, TotalCombinationCuts As Long, CombinationList() As Variant, CurrentFirstLocation As Long, Reverse As Boolean, IncList() As Integer
Dim good As Boolean, n As Integer, cList() As Double, cI As Long, LowN As Integer, Lcut As Integer, BlockDone As Boolean, j As Integer, Res As Boolean, reValue As Integer

TotalCuts = UBound(InputArr)
LowN = InputLD

If InputLD > 0 Then
    For a = LBound(InputArr) To UBound(InputArr)
        If InputArr(a) < LowN Then LowN = InputArr(a)
    Next a
    
    For a = LBound(InputArr) To UBound(InputArr)
       If InputArr(a) = LowN Then Hcut = Hcut + 1
    Next a
    
    Do
      sum = 0
      e = 0
      For b = LBound(InputArr) To UBound(InputArr)
        If InputArr(b) = LowN And e < Hcut Then
          sum = sum + InputArr(b)
          e = e + 1
        End If
      Next b
    If sum < InputLD Then
      Exit Do
Else:       Hcut = Hcut + -1
    End If
    Loop
Hcut = TotalCuts + -Hcut
End If

If InputLD = 0 Then Hcut = TotalCuts
i = 1
a = 0

Do Until a = Hcut + -1
  a = a + 1
  
  ReDim BlockArr(1 To a)
  y = TotalCuts + -a
For F = 1 To a
    y = y + 1
    BlockArr(F) = y
Next F

TotalCombinations = WorksheetFunction.Combin(TotalCuts, a)
IncDiff = a
TotalCombinationCuts = TotalCombinationCuts + (TotalCombinations * IncDiff) + (TotalCombinations * 6)
ReDim Preserve CombinationList(1 To TotalCombinationCuts)
First = True
g = 1
j = UBound(BlockArr)
Reverse = False
    
    n = i
    BlockDone = False
    Do Until BlockDone = True
                                                  
        For F = 1 To a
          
         If Reverse = True Then
             If Res = True Then
                reValue = BlockArr(j) + (a + -1)
                For c = LBound(BlockArr) To UBound(BlockArr)
                    If Not c = j Then
                      BlockArr(c) = reValue
                      reValue = reValue + -1
                    End If
                Next c
                Res = False
                g = TotalCuts
                If UBound(BlockArr) > 1 Then First = True
            End If
        End If
        
        If Reverse = False Then
            If Res = True Then
                reValue = BlockArr(j) + -(a + -1)
                For c = LBound(BlockArr) To UBound(BlockArr)
                    If Not c = j Then
                      BlockArr(c) = reValue
                      reValue = reValue + 1
                    End If
                Next c
                Res = False
                g = 1
                If UBound(BlockArr) > 1 Then First = True
            End If
        End If
        
        If First = True Then Exit For
                   
                  
          If BlockArr(j) = g Then
            If Reverse = True Then
               BlockDone = True
               Exit For
            End If
            If Reverse = False Then
             g = TotalCuts
             Reverse = True
             y = BlockArr(UBound(BlockArr))
             For b = LBound(BlockArr) To UBound(BlockArr)
                 BlockArr(b) = y
                 y = y + -1
             Next b
            End If
          End If

         If Reverse = True Then
           If BlockArr(F) = g Then
              g = g + -1
              Exit For
           End If
           If BlockArr(F) < g Then
                                BlockArr(F) = BlockArr(F) + 1
                                If Not j = F And BlockArr(F) = g Then g = g + -1
                                   If F = j Then
                                    If UBound(BlockArr) = 1 Then Res = True
                                    If UBound(BlockArr) > 1 Then
                                       If BlockArr(F + -1) + -BlockArr(F) > 1 Then Res = True
                                    End If
                                 End If
              Exit For
           End If
         End If
           
         If Reverse = False Then
           If BlockArr(F) = g Then
             g = g + 1
             Exit For
           End If
           If BlockArr(F) > g And Reverse = False Then
                                 BlockArr(F) = BlockArr(F) + -1
                                 If Not j = F And BlockArr(F) = g Then g = g + 1
                                 If F = j Then
                                    If UBound(BlockArr) = 1 Then Res = True
                                    If UBound(BlockArr) > 1 Then
                                       If -BlockArr(F + -1) + BlockArr(F) > 1 Then Res = True
                                    End If
                                 End If
             Exit For
           End If
         End If
         
        Next F
        
          First = False
          If BlockDone = True Then Exit Do
                        
                        ii = 1
                        ReDim IncList(1 To a)
                        For b = LBound(BlockArr) To UBound(BlockArr)
                            IncList(ii) = InputArr(BlockArr(b))
                            ii = ii + 1
                        Next b
           
           good = True
           sum = 0
           For F = LBound(IncList) To UBound(IncList)
               sum = sum + IncList(F)
           Next F
           
           cNumber = (sum * sum) + sum + IncDiff
           
            If k > 0 Then
             For F = LBound(cList) To UBound(cList)
                If cNumber = cList(F) Then good = False
             Next F
           End If
           
           If sum < InputLowD Then
              If -sum + InputLowD >= LowN Then good = False
           End If
           
           If sum > InputLD And InputLD > 0 Then good = False
           
           If good = True Then
              
              k = k + 1
              ReDim Preserve cList(1 To k)
              cList(k) = cNumber
           
         
           
                                                           CombinationList(i) = "Cuts"
                 i = i + 1
                                                           CurrentFirstLocation = i
                                                           CurrentLastLocation = i + (IncDiff + -1)
           
           
           For F = LBound(IncList) To UBound(IncList)
               CombinationList(i) = IncList(F)
                 i = i + 1
           Next F
           
                                                            CombinationList(i) = "Sum"
                 i = i + 1
                                                                                
           For F = CurrentFirstLocation To CurrentLastLocation
           
                                                            CombinationList(i) = CombinationList(i) + CombinationList(F)
           Next F
                i = i + 1
                                                            CombinationList(i) = "Number Of Cuts"

                i = i + 1
                                                            CombinationList(i) = IncDiff
                i = i + 1
            
                                                            CombinationList(i) = "end"
                                                            
                i = i + 1
                
                
        End If
        Loop
                
Loop

If TotalCuts > 1 Then
For F = LBound(CombinationList) To UBound(CombinationList)
    If F + 1 > UBound(CombinationList) Then Exit For
    If CombinationList(F + 1) = "" Then ReDim Preserve CombinationList(1 To F)
Next F
End If

If TotalCuts = 1 Then ReDim CombinationList(1 To 7)

ReDim Preserve CombinationList(1 To i + -1 + UBound(InputArr) + 6)

    CombinationList(i) = "Cuts"

sum = 0
For a = LBound(InputArr) To UBound(InputArr)
i = i + 1
    CombinationList(i) = InputArr(a)
    sum = sum + InputArr(a)
Next a

i = i + 1
    CombinationList(i) = "Sum"
i = i + 1
    CombinationList(i) = sum
i = i + 1
    CombinationList(i) = "Number Of Cuts"
i = i + 1
    CombinationList(i) = UBound(InputArr)
i = i + 1
    CombinationList(i) = "end"
    

                                FindAllCombinations = CombinationList
                                                            
End Function





