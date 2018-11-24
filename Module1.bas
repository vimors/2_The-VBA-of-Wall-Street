Attribute VB_Name = "Module1"
Dim tabla_final As New Collection
Dim element As cl_final
Sub get_total_volume()
Dim sh As Worksheet
Dim rw As Range
Dim element2 As cl_final
Dim TotalRange As Range
Set tabla_final = New Collection
Dim min_change As Double
Dim max_change As Double
Dim max_total As Double

Dim ticker_min_change As String
Dim ticker_max_change As String
Dim ticker_max_total As String

Dim indexRow As Integer
For Each sh In ThisWorkbook.Worksheets

    sh.Activate
    Set TotalRange = sh.UsedRange
    Set TotalRange = TotalRange.Offset(1, 0).Resize(TotalRange.Rows.Count - 1, TotalRange.Columns.Count)
    For Each rw In TotalRange.Rows
        Set element2 = searchInFinal(Cells(rw.Row, 1).Value)
        If Not element2 Is Nothing Then
            element2.Total_Volume = Cells(rw.Row, 7).Value + element2.Total_Volume
            If Cells(rw.Row, 2).Value > element2.Date_Close Then
                element2.Close_Stock = Cells(rw.Row, 6).Value
                element2.Date_Close = Cells(rw.Row, 2).Value
            End If
            If Cells(rw.Row, 2).Value < element2.Date_Open Then
                element2.Date_Open = Cells(rw.Row, 2).Value
                element2.Open_Stock = Cells(rw.Row, 3).Value
            End If
            element2.Yearly_Change = element2.Close_Stock - element2.Open_Stock
            If element2.Open_Stock <> 0 Then
                element2.Percent_Change = (element2.Yearly_Change * 100) / element2.Open_Stock
            Else
                element2.Percent_Change = 0
            End If
        Else
            Set element2 = New cl_final
            element2.Ticker = Cells(rw.Row, 1).Value
            element2.Total_Volume = Cells(rw.Row, 7).Value
            element2.Date_Open = Cells(rw.Row, 2).Value
            element2.Date_Close = Cells(rw.Row, 2).Value
            element2.Open_Stock = Cells(rw.Row, 3).Value
            element2.Close_Stock = Cells(rw.Row, 6).Value
            element2.Yearly_Change = element2.Close_Stock - element2.Open_Stock
            If element2.Open_Stock <> 0 Then
                element2.Percent_Change = (element2.Yearly_Change * 100) / element2.Open_Stock
            Else
                element2.Percent_Change = 0
            End If
            tabla_final.Add element2
            Set element2 = Nothing
        End If
        
    Next rw
    
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Volume"
    
    Cells(2, "O").Value = "Greatest % increase"
    Cells(3, "O").Value = "Greatest % Decrease"
    Cells(4, "O").Value = "Greatest total volume"
        
    Cells(1, "P").Value = "Ticker"
    
    indexRow = 2
    min_change = 0
    max_change = 0
    max_volume = 0
    Set element = Nothing
    
    For Each element In tabla_final
        Cells(indexRow, "I").Value = element.Ticker
        Cells(indexRow, "J").Value = element.Yearly_Change
        If element.Yearly_Change > 0 Then
            Cells(indexRow, "J").Interior.ColorIndex = 4
        Else
            Cells(indexRow, "J").Interior.ColorIndex = 6
        End If
        Cells(indexRow, "K").Value = element.Percent_Change
        Cells(indexRow, "L").Value = element.Total_Volume
        indexRow = indexRow + 1
        If element.Percent_Change < min_change Then
            ticker_min_change = element.Ticker
            min_change = element.Percent_Change
        End If
        If element.Percent_Change > max_change Then
            ticker_max_change = element.Ticker
            max_change = element.Percent_Change
        End If
        If element.Total_Volume > max_volume Then
            ticker_max_volume = element.Ticker
            max_volume = element.Total_Volume
        End If
    Next element
    
    Cells(2, "P").Value = ticker_max_change
    Cells(3, "P").Value = ticker_min_change
    Cells(4, "P").Value = ticker_max_volume
    
    Cells(2, "Q").Value = max_change
    Cells(3, "Q").Value = min_change
    Cells(4, "Q").Value = max_volume
    
    Set tabla_final = Nothing
Next sh

End Sub

Function searchInFinal(ByVal pTicker As String) As cl_final
    Dim element As cl_final
        For Each element In tabla_final
            If (element.Ticker = pTicker) Then
                Set searchInFinal = element
                Exit For
            End If
        Next element
        
End Function
Function setVolume(ByRef pVoloume As Long)
    Dim element As typeConsolidated
    For Each element In table_final
        If (element.Ticker = pElement.Ticker) Then
            element.Total_Volume = pvolume
            Exit For
        End If
    Next element
End Function


