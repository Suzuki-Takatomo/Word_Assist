Attribute VB_Name = "NewMacros"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
    Call Marker.next_sentence
End Sub
Sub macro1()
    With Selection.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = MillimetersToPoints(12.7)
        .BottomMargin = MillimetersToPoints(12.7)
        .LeftMargin = MillimetersToPoints(12.7)
        .RightMargin = MillimetersToPoints(12.7)
        .Gutter = MillimetersToPoints(0)
        .HeaderDistance = MillimetersToPoints(15)
        .FooterDistance = MillimetersToPoints(17.5)
        .PageWidth = MillimetersToPoints(210)
        .PageHeight = MillimetersToPoints(297)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
        .LinesPage = 36
        .LayoutMode = wdLayoutModeLineGrid
    End With
End Sub
