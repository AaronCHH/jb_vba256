Attribute VB_Name = "Module1"
Option Explicit

Sub 設定佈景主題()
    Dim tPath As String
    tPath = "C:\Program Files\Microsoft Office\Document Themes 12\"
    ActiveWorkbook.ApplyTheme (tPath & "Verve.thmx")
End Sub

Sub 變更配色()
    Dim cPath As String
    cPath = "C:\Program Files\Microsoft Office\Document Themes 12\Theme Colors\"
    ActiveWorkbook.Theme.ThemeColorScheme.Load (cPath & "Opulent.xml")
End Sub

Sub 變更字型()
    Dim fPath As String
    fPath = "C:\Program Files\Microsoft Office\Document Themes 12\Theme Fonts\"
    ActiveWorkbook.Theme.ThemeFontScheme.Load (fPath & "Equity.xml")
End Sub

Sub 變更效果()
    Dim ePath As String
    ePath = "C:\Program Files\Microsoft Office\Document Themes 12\Theme Effects\"
    ActiveWorkbook.Theme.ThemeEffectScheme.Load (ePath & "Verve.eftx")
End Sub




