Attribute VB_Name = "Module1"
Sub StockAnalysis():

Dim WorksheetName, Tickler, GIT, GDT, GVT As String
Dim Summary_Table_Row As Long
Dim SP, EP, Diff, GI, GD, GV As Double
Dim FirstRow As Boolean


For Each ws In Worksheets
        WorksheetName = ws.Name
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Range("i1").Value = "Tickler"
        ws.Range("j1").Value = "Yearly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("l1").Value = "Total Stock Volume"
        ws.Range("o2").Value = "Greatest % Increase"
        ws.Range("o3").Value = "Greatest % Decrease"
        ws.Range("o4").Value = "Greatest Total Volume"
        ws.Range("p1").Value = "Ticker"
        ws.Range("q1").Value = "Value"
     Summary_Table_Row = 2
     GI = 0
     GD = 0
     GV = 0

    For r = 2 To LastRow
    If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
    Tickler = ws.Cells(r, 1).Value
    EP = ws.Cells(r, 6).Value
    ws.Range("I" & Summary_Table_Row).Value = Tickler
    FirstRow = False
    Diff = (EP - SP)
    ws.Range("j" & Summary_Table_Row).Value = Diff
    ws.Range("j" & Summary_Table_Row).NumberFormat = "#,##0.00"

   If Diff < 0 Then
   ws.Range("j" & Summary_Table_Row).Interior.ColorIndex = 3
   ElseIf Diff > 0 Then
   ws.Range("j" & Summary_Table_Row).Interior.ColorIndex = 4
   End If

    ws.Range("k" & Summary_Table_Row).Value = (Diff / SP)
    ws.Range("k" & Summary_Table_Row).NumberFormat = "0.00%"
    MyVolume = MyVolume + ws.Range("G" & r).Value
    
    ws.Range("L" & Summary_Table_Row).Value = MyVolume
    'CHECK FOR GREATEST VALUES
    If MyVolume > GV Then
    GV = MyVolume
    GVT = Tickler
    End If
    
    If (Diff / SP) > GI Then
    GI = (Diff / SP)
    GIT = Tickler
    End If
    
    If (Diff / SP) < GD Then
    GD = (Diff / SP)
    GDT = Tickler
    End If

    Summary_Table_Row = Summary_Table_Row + 1
    MyVolume = 0

    Else
        If FirstRow = False Then
        SP = ws.Cells(r, 3).Value
        FirstRow = True
        End If
        MyVolume = MyVolume + ws.Range("G" & r).Value
   End If
   
   Next r
   
   'SET GREATEST
   ws.Range("P4").Value = GVT
   ws.Range("Q4").Value = GV
   
   ws.Range("P2").Value = GIT
   ws.Range("Q2").Value = GI
   ws.Range("Q2").NumberFormat = "0.00%"

   
   ws.Range("P3").Value = GDT
   ws.Range("Q3").Value = GD
   ws.Range("Q3").NumberFormat = "0.00%"

   
   Worksheets(WorksheetName).Columns("A:Q").AutoFit
   Next ws
End Sub
