Attribute VB_Name = "Module1"
Option Explicit

'This function opens and reads the text file and returns the number of credit cards
Public Function ReadFile(ByVal FName As String, CN() As String) As Integer
    Dim K As Integer
    
    K = 0
    
    Open FName For Input As #1
        Do While Not EOF(1)
            K = K + 1
            Input #1, CN(K)
        Loop
    Close #1
    
    ReadFile = K
End Function

'This function verifies if the credit card number is valid or not
Public Function InvalidCard(CN() As String, ByVal N As Integer) As Boolean
    Dim K As Integer
    Dim DblNum As Boolean   'Checks to see if the digit needs to be doubled
    Dim Digit As Integer
    Dim Sum As Integer
    
    'Assume credit card will be invalid
    InvalidCard = True
    'Initialize
    Sum = 0
    
    'If the credit card number has an even amount of digits, double the alternating digits
    'starting with the first digit
    If Len(CN(N)) Mod 2 = 0 Then
        DblNum = True
    'Otherwise, start doubling on the second digit
    Else
        DblNum = False
    End If
    
    For K = 1 To Len(CN(N))
        'Takes a digit from credit card number
        Digit = Val(Mid$(CN(N), K, 1))
        'If the digit needs to be doubled
        If DblNum = True Then
            Digit = Digit * 2
            'If the doubled digit is greater than 9, subtract the digit by 9
            If Digit > 9 Then
                Digit = Digit - 9
            End If
        End If
        'Add the digit to the running sum
        Sum = Sum + Digit
        'Alternate the doubling of digits
        DblNum = Not DblNum
    Next K
    
    'If the running sum is divisible by 10, then it is valid
    If Sum Mod 10 = 0 Then
        InvalidCard = False
    End If
End Function
