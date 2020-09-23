Attribute VB_Name = "Module1"
'
' Gamasort: from Joseph Gama
'
' This is a Radix Sort Exchange routine which will work only on
' unsigned integers (unlike a general sorting function which
' works based on comparisons).  Since its complexity overhead
' is pretty large, it is efficient on big array where the
' maximum number is bounded by a relatively small ceiling.
' E.g. a big array of bytes or shorts.
'
' The algorithm is pretty simple:
' 1) It will first find the maximum and minimum radices in order to see
'    if it can skip some radices.  Then it proceeds to order
'    from the numbers with the most-significant set-bit downwards
' 2) It leaves the unsorted segments of length == CUTOFF unsorted.
' 3) From higher bit to lower it will compare numbers
'    based on the bit being scanned.  This process is
'    improved by using 2 pivots that will move 0's to the
'    left and 1's to the right.
' 4) When finding the pivots it looks for the best case
'    of all 0's or 1's to skip more scanning.
' 5) A low overhead Insertion Sort finishes the job.
'
'
' *** To make it an O(n*log(n)) algorithm in which the 'n'
' is proportional to the largest number in the sorted array
' requiers commenting            Call InSort(t(), 1, up)
' and also commenting            (up-lo>CUTOFF)AND
' Which is clearly detailed in the code
'
' in this version, instead of including sort.h, SWAP and Insort were taken from
' "A library of internal sorting routines" by Ariel Faigon
' (http://www.yendor.com/programming/sort/)

Const CUTOFF = 15

Sub Swap(ByRef a As Long, ByRef b As Long)
'swap the contents of A and B
    a = a Xor b
    b = a Xor b
    a = a Xor b
End Sub

Function GT(x As Long, y As Long) As Long
'returns the greatest of 2 numbers
If x > y Then
    GT = x
Else
    GT = y
End If
End Function

Sub verify(a() As Long, i As Long, j As Long)
'to verify that the array was sorted properly
Dim x As Long
For x = i To j - 1
  If (a(x) > a(x + 1)) Then
  MsgBox "Error: Not sorted properly!!!"
  Exit Sub
  End If
Next
MsgBox "Array sorted properly!!!"
End Sub

Public Sub InSort(ByRef t() As Long, ByVal j As Long, ByVal n As Long)
'Insertion Sort
    Dim unsorted As Long, i As Long
    Dim nextValue As Long
    For unsorted = j To n
        nextValue = t(unsorted)      'next value to sort
        i = unsorted                 'index of next value
        If (i > j) Then
        Do While (t(i - 1) > nextValue)
            t(i) = t(i - 1)
            i = i - 1
            If i = 0 Then Exit Do    'special case - value is minimal
        Loop
        End If
        t(i) = nextValue             'insert new value
    Next unsorted
End Sub

Public Sub GamaSort(ByRef t() As Long, ByVal lo As Long, _
ByVal up As Long, ByVal lbit As Long, ByVal ubit As Long)
Dim temp, b, r, l As Long
    'b is the integer that is 2 to the power of the current bit being scanned
    'r is the index of the right pivot
    'l is the index of the left pivot
    If (up - lo > CUTOFF) And (ubit >= lbit) Then '(up - lo > CUTOFF) And
'(up-lo>CUTOFF)&& can be removed if k log n is desired for all cases
        'keep going only if
        'a)the sub-array is greater than the cutoff point
        'b)the current bit is not smaller than the lowest bit
        b = (2 ^ ubit) 'b=2^ubit
        r = up  '2 pivots
        l = lo
        While ((b And t(r)) = b) And (r > lo)
            r = r - 1
        Wend
        If Not ((r = lo) And ((b And t(r)) = b)) Then
            While ((b And t(l)) = 0) And (l < up)
            l = l + 1
            Wend
        End If
        If ((r = lo) And ((b And t(r)) = b)) Or ((l = up) And ((b And t(l)) = 0)) Then
            Call GamaSort(t(), lo, up, lbit, ubit - 1) 'if it"s all 0"cs or 1's, skip it
        Else
            'this is where the swapping occurs
            While l < r
                If ((b And t(l)) = b) And ((b And t(r)) = 0) Then
                    Call Swap(t(r), t(l))
                    While ((b And t(r)) = b) And (r > lo)
                        r = r - 1
                    Wend
                End If
            l = l + 1
            Wend
        If l >= r Then
            r = r + 1
        End If
        'recursivelly sort the 2 blocks corresponding to the sorted
        'sub-arrays but now for a lower radix
        If r - 1 > lo Then
            Call GamaSort(t(), lo, r - 1, lbit, ubit - 1) 'left
        End If
        If up > r Then
            Call GamaSort(t(), r, up, lbit, ubit - 1) 'right
        End If
    End If
    End If
End Sub

'
' Main version of gamasort (call this one)
'
Sub GamaSortStarter(ByRef t() As Long, ByVal up As Long)
Dim i As Long, ubit As Long, counter As Long, n As Long, lbit As Long, l As Long
ubit = 1
l = 1
        For counter = 0 To up
            n = n Or t(counter)
        Next
    If (n And 1) = 0 Then
        i = 1
        While (n And (2 ^ i)) = 0
        Wend
        lbit = i
    End If
    While (2 ^ ubit) <= n
        ubit = ubit + 1
    Wend
    ubit = ubit - 1
    Call GamaSort(t(), 1, up, lbit, ubit)
    Call InSort(t(), 1, up) 'if (up-lo>CUTOFF)AND was removed then this line should be removed as well
                        'if k log n is desired for all cases
End Sub

