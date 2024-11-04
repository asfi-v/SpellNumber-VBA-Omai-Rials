# SpellNumber-VBA-Omai-Rials

## Description
This repository contains an Excel VBA module that converts numbers to words. The `OMR` function can be used to transform numerical values into their corresponding English words of Omani Rials, which is particularly useful for financial documents and reports.

## Usage
1. Open your Excel workbook.
2. Press `ALT + F11` to open the VBA editor.
3. Click the Insert tab, and click Module.
   
   ![image](https://github.com/user-attachments/assets/50d06c7a-c9eb-4c5d-a74e-75c13009b2b8)
   
4. Copy the below VBA Code in Module.

   ![image](https://github.com/user-attachments/assets/5ba9c84d-d199-4faf-ae1a-ddb974d9e39f)
   
6. Use the `OMR` function in your Excel formulas.


## VBA Code:
```vba
Option Explicit

'Main Function

Function OMR(ByVal MyNumber)
    Dim Rial, Baizas, Temp
    Dim DecimalPlace, Count
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "
       MyNumber = Trim(Str(MyNumber))
        DecimalPlace = InStr(MyNumber, ".")
        If DecimalPlace > 0 Then
        Baizas = GetHundreds(Left(Mid(MyNumber, DecimalPlace + 1) & "000", 3))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    Count = 1
    Do While MyNumber <> ""
        Temp = GetHundreds(Right(MyNumber, 3))
        If Temp <> "" Then Rial = Temp & Place(Count) & Rial
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop
    Select Case Rial
        Case ""
            Rial = " Rial Omani Zero"
        Case "One"
            Rial = " Rial Omani One"
         Case Else
            Rial = " Rial Omani " & Rial
    End Select
    Select Case Baizas
        Case ""
            Baizas = ""
        Case "One"
            Baizas = " and Baiza's One"
              Case Else
            Baizas = " and " & " Baiza's " & Baizas
    End Select
    OMR = Rial & Baizas & " Only"
End Function
     

' Converts a number from 100-999 into text

Function GetHundreds(ByVal MyNumber)
    Dim Result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    ' Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
    End If
    ' Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTens(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If
    GetHundreds = Result
End Function
   

' Converts a number from 10 to 99 into text.
Function GetTens(TensText)
    Dim Result As String
    Result = ""           ' Null out the temporary function value.
    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...
        Select Case Val(TensText)
            Case 10: Result = "Ten"
            Case 11: Result = "Eleven"
            Case 12: Result = "Twelve"
            Case 13: Result = "Thirteen"
            Case 14: Result = "Fourteen"
            Case 15: Result = "Fifteen"
            Case 16: Result = "Sixteen"
            Case 17: Result = "Seventeen"
            Case 18: Result = "Eighteen"
            Case 19: Result = "Nineteen"
            Case Else
        End Select
    Else                                 ' If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: Result = "Twenty "
            Case 3: Result = "Thirty "
            Case 4: Result = "Forty "
            Case 5: Result = "Fifty "
            Case 6: Result = "Sixty "
            Case 7: Result = "Seventy "
            Case 8: Result = "Eighty "
            Case 9: Result = "Ninety "
            Case Else
        End Select
        Result = Result & GetDigit (Right(TensText, 1))  ' Retrieve ones place.
    End If
    GetTens = Result
End Function
    

' Converts a number from 1 to 9 into text.

Function GetDigit(Digit)
    Select Case Val(Digit)
        Case 1: GetDigit = "One"
        Case 2: GetDigit = "Two"
        Case 3: GetDigit = "Three"
        Case 4: GetDigit = "Four"
        Case 5: GetDigit = "Five"
        Case 6: GetDigit = "Six"
        Case 7: GetDigit = "Seven"
        Case 8: GetDigit = "Eight"
        Case 9: GetDigit = "Nine"
        Case Else: GetDigit = ""
    End Select
End Function
```

## Example
1. Type the formula =OMR(A1) into the cell where you want to display a written number, where A1 is the cell containing the number you want to convert. You can also manually type the value like =OMR(22.500)
2. Press Enter to confirm the formula.
3. Save your OMR function workbook.
4. _Excel cannot save a workbook with macro functions in the standard macro-free workbook format (.xlsx). If you click File > Save. A VB project dialog box opens. Click No._

## You can save your file as an **Excel Macro-Enabled Workbook (.xlsm)** to keep your file in its current format.
1. Click **File > Save As**.
2. Click the **Save as type** drop-down menu, and select **Excel Macro-Enabled Workbook**.
3. Click **Save**.
