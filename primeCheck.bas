Attribute VB_Name = "Module1"
Option Explicit
' option explicit giup cho kiem tra loi cu phap dat ten bien

Sub bangCuuChuong()
    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To 9
        For j = 1 To 9
            Cells(i, j) = i & "x" & j & "=" & i * j
        Next
    Next
    
    
End Sub
Sub ChanLe()
    Dim a As Integer
    
    a = InputBox("nhap a")
    If a Mod 2 = 0 Then
        MsgBox a & " la so chan"
    Else
        MsgBox a & " La so le"
    End If
    
    
    
    
End Sub

Sub primeCheck()
    Dim a As Integer
    Dim i As Integer
    Dim cnt As Integer
    a = InputBox("Nhap a")
    For i = 1 To a / 2
        If a Mod i = 0 Then
             cnt = cnt + 1
        Else
           
            cnt = 0
        End If
    Next
    If cnt > 0 Then
        MsgBox a & " khong phai la so nguyen to"
    Else
        MsgBox a & " la so nguyen to "
    End If
    
    

End Sub
