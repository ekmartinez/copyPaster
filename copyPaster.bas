Attribute VB_Name = "CopyPaster"

Sub CopyPaster()

    Dim sht As Worksheet

    Sheets("5720040 MAR FELICI").Range("D8:L27").Copy
        For Each sht In Worksheets
            sht.Range("D8:L27").PasteSpecial xlPasteAll
        Next

    Sheets("5720040 MAR FELICI").Range("C5").Copy
        For Each sht In Worksheets
            sht.Range("C5").PasteSpecial xlPasteAll
        Next

    Application.CutCopyMode = False
    
End Sub
