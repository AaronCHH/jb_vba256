Attribute VB_Name = "Module1"
Option Explicit

Sub �]�w�G���D�D()
    Dim tPath As String
    tPath = "C:\Program Files\Microsoft Office\Document Themes 12\"
    ActiveWorkbook.ApplyTheme (tPath & "Verve.thmx")
End Sub

Sub �ܧ�t��()
    Dim cPath As String
    cPath = "C:\Program Files\Microsoft Office\Document Themes 12\Theme Colors\"
    ActiveWorkbook.Theme.ThemeColorScheme.Load (cPath & "Opulent.xml")
End Sub

Sub �ܧ�r��()
    Dim fPath As String
    fPath = "C:\Program Files\Microsoft Office\Document Themes 12\Theme Fonts\"
    ActiveWorkbook.Theme.ThemeFontScheme.Load (fPath & "Equity.xml")
End Sub

Sub �ܧ�ĪG()
    Dim ePath As String
    ePath = "C:\Program Files\Microsoft Office\Document Themes 12\Theme Effects\"
    ActiveWorkbook.Theme.ThemeEffectScheme.Load (ePath & "Verve.eftx")
End Sub




