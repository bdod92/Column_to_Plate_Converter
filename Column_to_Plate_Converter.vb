'This method takes the raw QCO export data from the database and converts it into the standard plate format (16x24)
'The remaining columns are order dependent and should be: QCO, well, Conc_um, ReadOD, mw, e260.

Sub QCO_Sorter()

    '1. Enter the number of QCO plates that were generated from the COP
    plateNum = 6
    
    '2. Change ColumnNum to match the number of columns in your plate (QCO:24, FNT/QNT:3)
    columnNum = 24
    
    '3. Change plateRow to match the number of columns in your plate (QCO:16, FNT/QNT:2)
    plateRow = 16
    
    '4. Enter your desired spacing between plates
    plateSpace = 2
    
    '5. Counter for white space
    counter = 1
    
        For j = 1 To plateRow * plateNum
            
            'Reformat data from OMG LIMS raw data export. Single column to 384-well plate transform
            For i = 1 To columnNum
               Cells((j + counter), 7 + i) = Cells(columnNum * (j - 1) + i + 1, 4)
            Next i
            
            'Add space between plates after each 16th row
            If j Mod 16 = 0 Then
                counter = counter + plateSpace
            End If
     
        Next j
End Sub
