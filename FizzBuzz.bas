Attribute VB_Name = "FizzBuzz"
Sub FizzBuzz()

'Fizz = divisible by 3
'Buzz = divisible by 5
'FizzBuzz = divisible by 3 and 5

Dim i As Integer

For i = 1 To 100

    If i Mod 3 = 0 And i Mod 5 = 0 Then
    Debug.Print ("FizzBuzz")
    Else
    
        If i Mod 3 = 0 Then
        Debug.Print ("Fizz")
        Else
        
            If i Mod 5 = 0 Then
            Debug.Print ("Buzz")
            Else
            
            Debug.Print (i)
            
            End If
            
        End If
        
    End If

Next i

End Sub
