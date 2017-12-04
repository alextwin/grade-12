Attribute VB_Name = "Module1"
Option Explicit
Const MAX_WEEKDAY = 7
Const MAX_MONTH = 12

Public Function CheckMonth(ByVal M As Integer) As Integer
    Select Case M
        'If it is a 30 day month
        Case 4, 6, 9, 11
            CheckMonth = 30
        'If it is February
        Case 2
            CheckMonth = 28
        'Otherwise it is a 31 day month
        Case Else
            CheckMonth = 31
    End Select
End Function

Public Sub ChangeCalendar(ByVal Num As Integer, ByVal W As Integer, ByVal Grid As MSFlexGrid)
    Dim K As Integer
    Dim R As Integer        'Row number
    Dim Counter As Integer  'Used to place number in correct column
    Dim NumWeeks As Integer 'Number of weeks in a month
    Dim Start As Integer    'Used to calculate number of weeks
    
    'Initialize variables
    R = 1
    Counter = 1
    'Number of days + starting weekday
    Start = Num + W
    
    If Start Mod MAX_WEEKDAY = 0 Then
        'Number of weeks + 1
        '+1 for the grid
        NumWeeks = Start \ MAX_WEEKDAY + 1
    Else
        'Number of weeks + 2
        '+1 to round up numbers since it won't take up 1 full week
        '+1 for an extra row
        NumWeeks = Start \ MAX_WEEKDAY + 2
    End If
    
    'Insert rows
    Grid.Rows = NumWeeks
    
    'Delete previous text inside cells
    With Grid
        'Select rows and columns that are not fixed
        .Row = .FixedRows
        .Col = .FixedCols
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        'All selected cells will be affected by changes to the cell
        .FillStyle = flexFillRepeat
        'Clear text
        .Text = ""
        'Change background colour to gray
        .CellBackColor = &H808080
        'Only active cell will be affected by changes to the cell
        .FillStyle = flexFillSingle
    End With
    
    For K = 1 To NumWeeks - 1
        'Adjust the row height
        Grid.RowHeight(K) = 600
    Next K
    
    'Adjust the height of the calendar
    Grid.Height = Grid.RowHeight(0) + Grid.RowHeight(1) * (NumWeeks - 1) + 95
    
    'Insert numbers
    For K = 1 To Num
        Grid.Row = R
        Grid.Col = W + (K - Counter)
        Grid.Text = K
        'Centre numbers
        Grid.CellAlignment = flexAlignCenterCenter
        'Change the background colour of the cell to white
        Grid.CellBackColor = &HFFFFFF
        'Once column 6 has been reached (Saturday)
        If Grid.Col = 6 Then
            'Increment Counter and Row number
            Counter = Counter + 7
            R = R + 1
        End If
    Next K
End Sub

Public Sub Intialize(Grid As MSFlexGrid, ComboW As ComboBox, ComboM As ComboBox, W As Integer)
    Dim K As Integer
    
    'Initialize Weekday to 0
    W = 0
    'Set number of columns to number of weekdays
    Grid.Cols = MAX_WEEKDAY
    
    For K = 1 To MAX_WEEKDAY
        'Add Weekday name to combo box
        ComboW.AddItem WeekdayName(K)
        'Adjust column width
        Grid.ColWidth(K - 1) = 600
        Grid.Col = K - 1
        Grid.Row = 0
        'Add first 2 letters of the Weekday name to grid
        Grid.Text = Left$(WeekdayName(K), 2)
        'Centre Weekday names
        Grid.CellAlignment = flexAlignCenterCenter
    Next K
    
    'Adjust width of the grid
    Grid.Width = Grid.ColWidth(0) * 7 + 95
    
    For K = 1 To MAX_MONTH
        'Add Month name to combo box
        ComboM.AddItem MonthName(K)
    Next K
    
    'Set list index to 0
    ComboM.ListIndex = 0
    ComboW.ListIndex = 0
End Sub
