
# charts

### Resize All Charts on a Worksheet
    'Step 1: Declare your variables
    Dim i As Integer
    'Step 2: Start Looping through all the charts
    For i = 1 To ActiveSheet.ChartObjects.Count
    'Step 3: Activate each chart and size
    With ActiveSheet.ChartObjects(i)
    .Width = 300
    .Height = 200
    End With
    'Step 4: Increment to move to next chart
    Next i
