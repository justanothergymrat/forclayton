Sub main()

'####################
'####              define stuff
'####################
Dim sheetname As String

'####################
'####            cycle through
'####################

For Each Sheet In Worksheets

sheetname = Sheet.Name                      '#pull name from sheet
Worksheets(sheetname).Activate            '#prop it up

namepopups                                          '#do things to it
comparestats
formatstuff

Next Sheet

End Sub

Sub namepopups()

'####################
'####              define stuff
'####################
Dim names As String
Dim names2 As String
Dim cols As Integer
Dim tally As Integer
Dim value As Long
Dim lastrow As Long

Dim yearopen As Single
Dim yearclose As Single

'####################
'####            set headers
'####################
Cells(1, 9).value = "ticker"
Cells(1, 10).value = "yearly change"
Cells(1, 11).value = "% change"
Cells(1, 12).value = "ticker volume"
Cells(1, 13).value = "days of data"
Cells(1, 14).value = "year open"
Cells(1, 15).value = "year close"

'####################
'####            set variables
'####################

names = blank
tally = 0
cols = 1
value = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'####################
'#### cycle through names
'####       & find duplicates
'####################

For i = 1 To lastrow
'For i = 1 To 10
    names = Cells(i, 1).value
          
    If names <> "<ticker>" And names <> "" Then         '# ignore headers
            
    If names = names2 Then                                           '# increment current name
            tally = tally + 1
            Cells(cols, 13).value = tally
            Cells(cols, 12).value = Cells(cols, 12).value + Cells(i, 7).value
            'Cells(cols,12).value = Cells(i,7).value + value
            ' value = Cells(cols, 12).value      <-- causes overflow cause dumb.
            yearclose = Application.WorksheetFunction.Round((Cells(i, 6).value), 2)
                        If i = lastrow Then
                            Cells(cols, 13).value = tally
                            Cells(cols, 15).value = yearclose
                            Cells(cols, 10).value = yearclose - yearopen
                            If yearopen <> 0 Then                                   '#cannot divide by 0
                                Cells(cols, 11).value = (yearclose - yearopen) / yearopen
                            End If
                            If Cells(cols, 11).value >= 0 Then
                                Cells(cols, 11).Interior.ColorIndex = 4
                            Else
                                Cells(cols, 11).Interior.ColorIndex = 3
                            End If
                End If
    Else                                                                            '# unique name
            If tally = 0 Then                                                  '# first runs only
                cols = cols + 1
                tally = 1
                    Cells(cols, 13).value = tally
   '             Cells(i, 1).value = names
   '                 Cells(cols, 9).value = names
   '             names2 = Cells(i, 1).value
   '             value = Cells(i, 7).value
   '                 Cells(cols, 12).value = value
   '             yearopen = Application.WorksheetFunction.Round((Cells(i, 3).value), 2)
   '                 Cells(cols, 14).value = yearopen
   '             yearclose = Application.WorksheetFunction.Round((Cells(i, 6).value), 2)
   '                 Cells(cols, 15).value = yearclose
            Else                                                                       '# finish calc for current line
                Cells(cols, 13).value = tally
                Cells(cols, 15).value = yearclose
                Cells(cols, 10).value = yearclose - yearopen
                   ' If yearopen = 0 Then
                    '    Cells(cols, 11).value = "NaN"
                    'Else
                    '    Cells(cols, 11).value = (yearclose - yearopen) / yearopen
                    'End If
                    If yearopen <> 0 Then                                   '#cannot divide by 0
                        Cells(cols, 11).value = (yearclose - yearopen) / yearopen
                    End If
                    If Cells(cols, 11).value >= 0 Then
                        Cells(cols, 11).Interior.ColorIndex = 4
                    Else
                        Cells(cols, 11).Interior.ColorIndex = 3
                    End If
                tally = 1
                cols = cols + 1
   '                 Cells(cols, 9).value = names
   '             names2 = Cells(i, 1).value
   '             value = Cells(i, 7).value
   '                 Cells(cols, 12).value = value
   '             yearopen = Application.WorksheetFunction.Round((Cells(i, 3).value), 2)
   '                 Cells(cols, 14).value = yearopen
   '             yearclose = Application.WorksheetFunction.Round((Cells(i, 6).value), 2)
   '                 Cells(cols, 15).value = yearclose
            End If                                                                    '# common stuff read per line
                    Cells(i, 1).value = names
                    Cells(cols, 9).value = names
                names2 = Cells(i, 1).value
                value = Cells(i, 7).value
                    Cells(cols, 12).value = value
                yearopen = Application.WorksheetFunction.Round((Cells(i, 3).value), 2)
                    Cells(cols, 14).value = yearopen
                yearclose = Application.WorksheetFunction.Round((Cells(i, 6).value), 2)
                    Cells(cols, 15).value = yearclose
        End If
        End If
        
Next

End Sub

Sub comparestats()
'####################
'####              define stuff
'####################
Dim lastrow As Long

'####################
'####            set headers
'####################
Cells(1, 18).value = "ticker"
Cells(1, 19).value = "value"
Cells(2, 17).value = "big increase"
Cells(3, 17).value = "big decrease"
Cells(4, 17).value = "big volume"

'####################
'####            set variables
'####################
lastrow = Cells(Rows.Count, 12).End(xlUp).Row

'####################
'####            cycle through
'####################

For i = 2 To lastrow

'####################
'####            isolate big volume
'####################
'If Cells(i, 12).value = starting Then
'    Cells(4, 19).value = Cells(i, 12).value
'    Cells(4, 18).value = Cells(i, 9).value
If Cells(i, 12).value > Cells(4, 19).value Then
    Cells(4, 19).value = Cells(i, 12).value
    Cells(4, 18).value = Cells(i, 9).value
End If
'####################
'####            isolate big decrease
'####################
'If Cells(i, 10).value < starting Then
'    Cells(3, 19).value = Cells(i, 10).value
'    Cells(3, 18).value = Cells(i, 9).value
If Cells(i, 10).value < Cells(3, 19).value Then
    Cells(3, 19).value = WorksheetFunction.Round(Cells(i, 10).value, 2)
    Cells(3, 18).value = Cells(i, 9).value
End If
'####################
'####            isolate big increase
'####################
'If Cells(i, 12).value = starting Then
'    Cells(2, 19).value = WorksheetFunction.Round((Cells(i, 10).value), 2)
'    Cells(2, 18).value = Cells(i, 9).value
If Cells(i, 10).value > Cells(2, 19).value Then
    Cells(2, 19).value = WorksheetFunction.Round(Cells(i, 10).value, 2)
    Cells(2, 18).value = Cells(i, 9).value
End If


Next

End Sub

Sub formatstuff()
'####################
'####              define stuff
'####################
Dim lastrow As Long

'####################
'####             make pretty
'####################
lastrow = Cells(Rows.Count, 10).End(xlUp).Row               '#yearly change
Range("J2:J" & lastrow).NumberFormat = "0.00"
lastrow = Cells(Rows.Count, 14).End(xlUp).Row               '#year open
Range("n2:n" & lastrow).NumberFormat = "0.00"
lastrow = Cells(Rows.Count, 15).End(xlUp).Row               '#year close
Range("o2:o" & lastrow).NumberFormat = "0.00"
lastrow = Cells(Rows.Count, 11).End(xlUp).Row               '#% change
Range("k2:k" & lastrow).NumberFormat = "0.00%"

End Sub
