Attribute VB_Name = "Module1"
Sub Nike_test()

'Declaring variables
Dim LRM As Long
Dim LRR As Long
Dim originOnhand As String
Dim storeNum As String
Dim lookup_wgt As Double
Dim tempCost As Double
Dim tempMax As Double
Dim tempMin As Double
Dim freightCost As Double
Dim found As Boolean
Dim aptDate As Date
Dim aptTime As Date
Dim fuelRate As Double
Dim fuelCharge As Double
Dim totalCost As Double
Dim i As Long, j As Long
Dim wgtBrk As Integer
Dim Comment As String
Dim packFlatRate As Double
Dim totalPackWeight As Double
Dim zone As String

' Fetch the fuel rate from 'Rates' worksheet (cell A:10)
fuelRate = shMain.Range("B2")

' Find last row in both sheets
LRM = shMain.Cells(Rows.Count, 1).End(xlUp).Row 'Last row main
LRR = shRates.Cells(Rows.Count, 1).End(xlUp).Row 'Last row rates

If LRM <= 5 Then
    MsgBox "No data entered.  Please enter and run again."
    Exit Sub
End If

'Get total weight of the appt
For i = 6 To LRM
    totalWgt = totalWgt + shMain.Cells(i, 5)
Next i

aptDate = shMain.Cells(6, 8) 'APT Date column in main
aptTime = shMain.Cells(6, 9) 'APT Time column in main

'If Weekend
If Weekday(aptDate, vbMonday) >= 6 And special = "" Then
    special = "weekend"
    freightCost = 310
    ' fuelCharge = 0 'flat rate, no fuel
     Comment = "Weekend rate of $" & freightCost & " was applied."
     'shMain.Cells(i, 10).Value = freightCost
'if early or late
ElseIf (Hour(aptTime) >= 18 Or Hour(aptTime) < 8) And special = "" Then
    special = "afterhours"
      ' PM delivery (after 6 PM)
      freightCost = 250
        'fuelCharge = 0 'flat rate, no fuel
      Comment = "AM/PM rate of $" & freightCost & " was applied."
      'shMain.Cells(i, 10).Value = freightCost
Else
    special = ""
End If

'calculate if special
If special <> "" Then
    For i = 6 To LRM
        shMain.Cells(i, 12) = Round((shMain.Cells(i, 5) / totalWgt) * freightCost, 2)
        shMain.Cells(i, 13) = Comment
    Next i
    For i = 6 To LRM
        totalCost = totalCost + shMain.Cells(i, 12)
    Next i
    If totalCost <> freightCost Then
        shMain.Cells(LRM, 12) = shMain.Cells(LRM, 12) + (freightCost - totalCost)
    End If

'End If

''Calculate if not special
'If special = "" Then
'
'
'End If

Else
'Standard Rate Calculation
For i = 6 To LRM
    originOnhand = shMain.Cells(i, 1) 'Origin Onhand column in main
    storeNum = shMain.Cells(i, 4) 'Store Number column in main
    lookup_wgt = shMain.Cells(i, 5) 'LBS column in main
    aptDate = shMain.Cells(i, 8) 'APT Date column in main
    aptTime = shMain.Cells(i, 9) 'APT Time column in main

    'Initialize values
        tempCost = 0
        tempMax = 0
        tempMin = 0
        freightCost = 0
        fuelCharge = 0
        totalCost = 0
        Comment = ""
        found = False
'
'      ' Check if it's a weekend or after 6 PM
'    If Weekday(aptDate, vbMonday) >= 6 Then
'            ' Weekend (Saturday or Sunday)
'            freightCost = 310
'           ' fuelCharge = 0 'flat rate, no fuel
'            Comment = "Weekend rate applied."
'            'shMain.Cells(i, 10).Value = freightCost
'        ElseIf Hour(aptTime) >= 18 Or Hour(aptTime) < 8 Then
'            ' PM delivery (after 6 PM)
'            freightCost = 250
'          '  fuelCharge = 0 'flat rate, no fuel
'            Comment = "AM/PM rate was applied."
'            'shMain.Cells(i, 10).Value = freightCost
'        Else
'            ' Normal rate logic applies
'            packFlatRate = 0 'Not a flat rate
'        End If
'
'    ' Calculate total weight for the pack if flat rate applies
'    If packFlatRate > 0 Then
'        totalPackWeight = 0
'
'        'Sum up total LBS for the pack
'        For j = 6 To LRM
'            If shMain.Cells(j, 8).Value = aptDate And shMain.Cells(j, 9).Value = aptTime Then
'                totalPackWeight = totalPackWeight + shMain.Cells(j, 5).Value 'Sum of LBS in pack
'            End If
'        Next j
'
'        'Calculate proportional cost for each shipment in the pack
'        If totalPackWeight > 0 Then
'            totalCost = Round((lookup_wgt / totalPackWeight) * packFlatRate, 2)
'            fuelCharge = 0 'No fuel charge
'        End If
'
'        'Populate values in Main
'        shMain.Cells(i, 10).Value = freightCost
'        shMain.Cells(i, 11).Value = fuelCharge
'        shMain.Cells(i, 12).Value = totalCost
'        shMain.Cells(i, 13).Value = Comment
'
'    Else
    'Standard Rates Calculation:
    For j = 2 To LRR
        tempStore = shRates.Cells(j, 2).Value

        If tempStore = storeNum Then
        found = True
        zone = shRates.Cells(j, 5).Value ' Get the zone from Rates sheet

        'Rates Calculation

    'Weight break logic
         If lookup_wgt >= 5000 Then wgtBrk = 13
         If lookup_wgt >= 4000 And lookup_wgt < 5000 Then wgtBrk = 12
        If lookup_wgt >= 3000 And lookup_wgt < 4000 Then wgtBrk = 11
         If lookup_wgt >= 2000 And lookup_wgt < 3000 Then wgtBrk = 10
         If lookup_wgt >= 1000 And lookup_wgt < 2000 Then wgtBrk = 9
        If lookup_wgt >= 500 And lookup_wgt < 1000 Then wgtBrk = 8
        If lookup_wgt >= 100 And lookup_wgt < 500 Then wgtBrk = 7

    'Setting variables for Max/Min & Final Cost
        tempCost = lookup_wgt * shRates.Cells(j, wgtBrk)
        tempMax = shRates.Cells(j, 14).Value
        tempMin = shRates.Cells(j, 6).Value

     'Determines whether or not Max/Min value or rate calculation is used based on tempCost
        If tempCost >= tempMax Then
            freightCost = tempMax
            Comment = "Max rate applied. Zone: " & zone
        ElseIf tempCost <= tempMin Then
            freightCost = tempMin
            Comment = "Min rate applied. Zone: " & zone
        Else
            freightCost = tempCost
            Comment = "Zone: " & zone 'Default comment for when no max/min or weekend/pm rate is used
        End If

    'Populates freightCost values into 'Cost' column in Main
        shMain.Cells(i, 10).Value = freightCost
        Exit For
       End If
    Next j
'End If
'
'    ' If no matching store is found
'      '  If found = False Then
'       '  shMain.Cells(i, 13).Value = "Store Not Found"
'        ' Else
'        'End If
'
'
'        ' Calculate the fuel charge and total cost
'        ' fuelCharge = Round(fuelRate * freightCost, 2)
        If freightCost = 310 Or freightCost = 250 Then 'In the case of a flat rate
            fuelCharge = 0
        Else
            'Fuel charge calculation for non-flat rates
            fuelCharge = Round(fuelRate * freightCost, 2)
        End If

        totalCost = Round(freightCost + fuelCharge, 2)

        ' Populate fuel charge/total cost in Main (columns K & L)
        shMain.Cells(i, 11).Value = fuelCharge
        shMain.Cells(i, 12).Value = totalCost

        ' Populate the Comments section in Main (Column 13)
        shMain.Cells(i, 13).Value = Comment
'    'Look up Store Number in the "Rates" sheet based on Origin Onhand
'        'storeNumber = Application.VLookup(originOnhand, shRates.Range("A:C"), 3, False)
'
'    'If found, display the Store Number
'        'If Not IsError(storeNumber) Then
'            'shMain.Cells(i,4).Value = storeNumber
'        'Else
'            'shMain.Cells(i,4).Value = "Not Found"
'
'       ' End If
'
    Next i
    End If
        
End Sub

Sub SearchFileNumber()

    ' Declare variables
    Dim LRM As Long
    Dim fileNum As String
    Dim originOnhand As String
    Dim storeNum As String
    Dim lbs As Double
    Dim aptDate As Date
    Dim aptTime As Date
    Dim i As Long
    Dim outputRow As Long
    Dim found As Boolean
    Dim freightCost As Double

    ' Get the file number from cell B1 in the Search sheet
    fileNum = Sheets("Search").Range("B1").Value

    ' Set outputRow to start at row 2 for output
    outputRow = 2

    ' Find the last row in the Main sheet directly
    LRM = Sheets("Main").Cells(Sheets("Main").Rows.Count, 1).End(xlUp).Row

    ' Initialize found as False
    found = False

    ' Loop through the Main sheet to find the matching file number
    For i = 2 To LRM
        If Sheets("Main").Cells(i, 3).Value = fileNum Then ' SLC File - J#
            ' If the file is found, populate the values
            found = True

            ' Get values from the Main sheet
            originOnhand = Sheets("Main").Cells(i, 1).Value ' Origin Onhand column in Main
            storeNum = Sheets("Main").Cells(i, 4).Value ' Store Number column in Main
            lbs = Sheets("Main").Cells(i, 5).Value ' LBS column in Main
            aptDate = Sheets("Main").Cells(i, 9).Value ' APT Date column in Main
            aptTime = Sheets("Main").Cells(i, 10).Value ' APT Time column in Main
            freightCost = Sheets("Main").Cells(i, 10).Value ' Cost column in Main

            ' Populate values in the Search sheet
            Sheets("Search").Cells(outputRow, 2).Value = originOnhand ' Origin Onhand
            Sheets("Search").Cells(outputRow + 1, 2).Value = storeNum ' Store Number
            Sheets("Search").Cells(outputRow + 2, 2).Value = lbs ' LBS
            Sheets("Search").Cells(outputRow + 3, 2).Value = aptDate ' APT Date
            Sheets("Search").Cells(outputRow + 4, 2).Value = aptTime ' APT Time
            Sheets("Search").Cells(outputRow + 5, 2).Value = freightCost 'Cost

            ' Exit after finding the first match
            Exit For
        End If
    Next i

    ' If no match is found, notify the user
    If found = False Then
        MsgBox "File Number not found.", vbExclamation
    End If

End Sub

Sub clearSheet()

Range("A6:M1000").ClearContents

End Sub

Function find_fuel_cost(ByVal cost As Double) As Double
 
 
last_row = shSLAC.Range("A" & Rows.Count).End(xlUp).Row
found_cost = 0#
found = False
 
For a = 2 To last_row
    min_cost = shSLAC.Range("A" & a)
    max_cost = shSLAC.Range("C" & a)
    If (cost <= max_cost) And (cost >= min_cost) Then
        found = True
        found_cost = shSLAC.Range("E" & a).Value
        find_fuel_cost = found_cost
        Exit Function
    End If
Next a
 
If found = False Then
    find_fuel_cost = found_cost
End If
 
End Function
