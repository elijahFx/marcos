Sub CalculatePenaltyAndReplace()
    Dim selectedText As String
    Dim daysDifference As Long
    Dim penaltyAmount As Double
    Dim newText As String
    Dim contractAmount As Double
    Dim penaltyRate As Double
    Dim startDate As Date, endDate As Date
    
    penaltyRate = 0.01
    selectedText = Selection.Text
    
    
    Dim dateRegex As Object
    Set dateRegex = CreateObject("VBScript.RegExp")
    dateRegex.Global = False
    dateRegex.IgnoreCase = True
    dateRegex.Pattern = "\b(\d{2}\.\d{2}\.\d{4})\s*по\s*(\d{2}\.\d{2}\.\d{4})\b"
    
    If dateRegex.Test(selectedText) Then
        Dim dateMatches As Object
        Set dateMatches = dateRegex.Execute(selectedText)
        Dim startDateStr As String, endDateStr As String
        startDateStr = dateMatches(0).SubMatches(0)
        endDateStr = dateMatches(0).SubMatches(1)
        
       
        startDate = DateSerial(CInt(Right(startDateStr, 4)), CInt(Mid(startDateStr, 4, 2)), CInt(Left(startDateStr, 2)))
        endDate = DateSerial(CInt(Right(endDateStr, 4)), CInt(Mid(endDateStr, 4, 2)), CInt(Left(endDateStr, 2)))
        
        daysDifference = DateDiff("d", startDate, endDate) + 1
    Else
        MsgBox "Не найдены даты в формате 'dd.mm.yyyy по dd.mm.yyyy'."
        Exit Sub
    End If
    
    
    newText = Replace(selectedText, "X", daysDifference)
    
    
    Dim amountRegex As Object
    Set amountRegex = CreateObject("VBScript.RegExp")
    amountRegex.Global = True
    amountRegex.IgnoreCase = True
    amountRegex.Pattern = "\d{1,3}(?:\s?\d{3})*,\d{2}"
    
    Dim contractStr As String
    If amountRegex.Test(newText) Then
        Dim amountMatches As Object
        Set amountMatches = amountRegex.Execute(newText)
        
        contractStr = Replace(amountMatches(0).Value, " ", "")
        contractStr = Replace(contractStr, Chr(160), "")
        
        If Not IsNumeric(contractStr) Then
            MsgBox "Ошибка: '" & contractStr & "' не является числовым значением."
            Exit Sub
        End If
        
        contractAmount = CDbl(contractStr)
    Else
        MsgBox "Не найдена сумма договора в тексте."
        Exit Sub
    End If
    
   
    penaltyAmount = contractAmount * penaltyRate * daysDifference
    
  
    Dim penaltyFormatted As String
      penaltyFormatted = Replace(Format(penaltyAmount, "###0.00"), ".", ",")
    
  
    newText = Replace(newText, "Y", penaltyFormatted)
    
   
    Selection.Text = newText
End Sub

