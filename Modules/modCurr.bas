Attribute VB_Name = "modCurr"
    Function CurrencyToText(curValue As Currency) As String
    Static Ones(10) As String
    Static Teens(10) As String
    Static Tens(10) As String
    Static Thousands(3) As String
    Dim i As Integer, nPosition As Integer
    Dim nNumber As Integer, nStars As Integer
    Dim bZeroValue As Boolean
    Dim stResult As String, stTemp As String, stStars As String
    Dim stBuffer As String


    If curValue > 999999.99 Then
        MsgBox "The limit of this Function is 999999.99", vbExclamation, "Out of range error.."
        Exit Function
    End If
    Ones(0) = "Zero"
    Ones(1) = "One"
    Ones(2) = "Two"
    Ones(3) = "Three"
    Ones(4) = "Four"
    Ones(5) = "Five"
    Ones(6) = "Six"
    Ones(7) = "Seven"
    Ones(8) = "Eight"
    Ones(9) = "Nine"
    Teens(0) = "Ten"
    Teens(1) = "Eleven"
    Teens(2) = "Twelve"
    Teens(3) = "Thirteen"
    Teens(4) = "Fourteen"
    Teens(5) = "Fifteen"
    Teens(6) = "Sixteen"
    Teens(7) = "Seventeen"
    Teens(8) = "Eighteen"
    Teens(9) = "Nineteen"
    Tens(0) = ""
    Tens(1) = "Ten"
    Tens(2) = "Twenty"
    Tens(3) = "Thirty"
    Tens(4) = "Forty"
    Tens(5) = "Fifty"
    Tens(6) = "Sixty"
    Tens(7) = "Seventy"
    Tens(8) = "Eighty"
    Tens(9) = "Ninty"
    Thousands(0) = ""
    Thousands(1) = "Thousand"
    
    'Set the cents portion of the string
    stResult = "and " & Format((curValue - Int(curValue)) * 100, "00") & "/100"
    'Convert the dollar portion to a string
    stTemp = CStr(Int(curValue))
    'parse through string(Dollar ammount)


    For i = Len(stTemp) To 1 Step -1
        'Grab the value of this digit
        nNumber = Val(Mid(stTemp, i, 1))
        'Check the position(column) of this digi
        '     t
        'Ones, Tens, or Hundereds
        nPosition = (Len(stTemp) - i) + 1


        Select Case (nPosition Mod 3)
            Case 1 'Ones position
            bZeroValue = False


            If i = 1 Then
                stBuffer = Ones(nNumber) & " "
            ElseIf Mid(stTemp, i - 1, 1) = "1" Then
                stBuffer = Teens(nNumber) & " "
                i = i - 1 'Skip tens position
            ElseIf nNumber > 0 Then
                stBuffer = Ones(nNumber) & " "
            Else
                bZeroValue = True


                If i > 1 Then


                    If Mid(stTemp, i - 1, 1) <> "0" Then
                        bZeroValue = False
                    End If
                End If


                If i > 2 Then


                    If Mid(stTemp, i - 2, 1) <> "0" Then
                        bZeroValue = False
                    End If
                End If
                stBuffer = ""
            End If


            If bZeroValue = False And nPosition > 1 Then
                stBuffer = stBuffer & Thousands(nPosition / 3) & " "
            End If
            stResult = stBuffer & stResult
            Case 2 'Tens position
            'Numbers like twenty-five need to be hyp
            '     henated. So......
            'Check if the digit has a value other th
            '     an 0 AND check the next
            'digit to see if it has a value other th
            '     an 0
            'if both are true add the hyphen


            If nNumber > 0 And Val(Mid(stTemp, i + 1, 1)) = 0 Then
                stResult = Tens(nNumber) & " " & stResult
            ElseIf nNumber > 0 And Val(Mid(stTemp, i + 1, 1)) > 0 Then
                stResult = Tens(nNumber) & "-" & stResult
            End If
            Case 0 'Hundreds position


            If nNumber > 0 Then
                stResult = Ones(nNumber) & " Hundred " & stResult
            End If
        End Select
Next i


If Len(stResult) > 0 Then
    stResult = UCase(Left(stResult, 1)) & Mid(stResult, 2)
End If
nStars = 125 - Len(stResult)


For i = 0 To nStars
    stStars = stStars + "*"
Next
CurrencyToText = stResult & " " & stStars
End Function

