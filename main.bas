Sub FetchJSONToSelectedCell()
    Dim url As String
    Dim response As String
    Dim oCell As Object
    Dim symbolCell As Object
    Dim symbolSheet As Object
    Dim symbolRange As Object
    Dim lastRow As Integer
    Dim i As Integer

    ' Define the API endpoint for NAV data
    Dim baseURL As String
    baseURL = "https://query1.finance.yahoo.com/v8/finance/chart/"

    ' Get the symbol sheet
    Set symbolSheet = ThisComponent.Sheets.getByName("rawdaat")

    ' Initialize lastRow to a starting value
    lastRow = 1

    ' Loop through each cell in column B to find the last non-empty row
    Do While symbolSheet.getCellByPosition(1, lastRow).getString() <> ""
        lastRow = lastRow + 1
    Loop

    ' Loop through each cell in column B
    For i = 1 To lastRow - 1 ' Assuming data starts from row 1
        ' Get the symbol from the current cell in column B
        Set symbolCell = symbolSheet.getCellByPosition(1, i)

        ' Get the symbol
        Dim symbol As String
        symbol = symbolCell.getString()

        ' Construct the URL
        url = baseURL & symbol

        ' Make the GET request to fetch data from API
        response = GetHTTP(url)

        ' Get the corresponding cell in column C
        Set oCell = symbolSheet.getCellByPosition(2, i)

        ' Paste the JSON response into the corresponding cell in column C
        oCell.setString(response)
    Next i
End Sub
