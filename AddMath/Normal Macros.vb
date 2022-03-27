Sub AddDate()
'
' AddDates Macro
'
'
    Selection.InsertDateTime DateTimeFormat:="dddd dd MMMM yyyy", _
        InsertAsField:=False, DateLanguage:=wdHebrew, CalendarType:= _
        wdCalendarHebrew, InsertAsFullWidth:=False
End Sub
Sub CopyParagraph()
'
' CopyParagraph Macro
'
'
    Selection.MoveUp Unit:=wdParagraph, Count:=1
    Selection.MoveDown Unit:=wdParagraph, Count:=1, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
End Sub

Sub InitializeCustomMath()
    Application.CustomizationContext = ThisDocument.AttachedTemplate
    KeyBindings.Add wdKeyCategoryMacro, "Normal.NewMacros.AddCustomMathCs", 220
    KeyBindings.Add wdKeyCategoryMacro, "Normal.NewMacros.AddSlash", BuildKeyCode(wdKeyAlt, 220)
End Sub
Sub AddSlash()
    Selection.TypeText "\"
End Sub
Sub AddCustomMathCs()
    If Selection.OMaths.Count > 0 Then
        On Error GoTo Error
        AppActivate "AddCustomMath"
    Else
        Selection.TypeText "\"
    End If
    Exit Sub
Error:
    Selection.TypeText "\"
End Sub

Sub AddCustomMath()
    AddMath.Show
End Sub

Sub SaveOnDesktop()
'
' SaveOnDesktop Macro
'
'
    Dim FileName  As String
    FileName = "C:\Users\Oryan\Desktop\" & ActiveDocument.BuiltInDocumentProperties("title") & ".pdf"
    ChangeFileOpenDirectory "C:\Users\Oryan\Desktop\"
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        FileName, ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=True, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False
End Sub



' in ThisDocument
Private Sub Document_New()
    InitializeCustomMath()
End Sub

Private Sub Document_Open()
    InitializeCustomMath()
End Sub
