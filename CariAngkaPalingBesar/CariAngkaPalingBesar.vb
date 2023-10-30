Module CariAngkaPalingBesar
    'Buka CariAngkaPalingBesar.vbproj
    Sub Main()
        Dim a As Long, b As Long, c As Long
        a = InputBox("Masuk nilai ke-1:")
        b = InputBox("Masuk nilai ke-2:")
        c = InputBox("Masuk nilai ke-3:")
        Debug.Print("Nilai a = " & a)
        Debug.Print("Nilai b = " & b)
        Debug.Print("Nilai c = " & c)
        If (a > b) And (a > c) Then
            Debug.Print("Nilai terbesar = " & a)
        ElseIf b > c Then
            Debug.Print("Nilai terbesar = " & b)
        Else
            Debug.Print("Nilai terbesar = " & c)
        End If
    End Sub

End Module
