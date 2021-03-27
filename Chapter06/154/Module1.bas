Attribute VB_Name = "Module1"
Option Explicit

Sub ¥Ø¿ý¦Cªí()
    Dim mFolder As String, sFolder As String, i As String
    mFolder = "C:\ExcelVBA\"
    sFolder = Dir(mFolder, vbDirectory)
    i = 3
    Do While sFolder <> ""
        If sFolder <> "." And sFolder <> ".." Then
            If GetAttr(mFolder & sFolder) And vbDirectory Then
                Cells(i, 1).Value = sFolder
                i = i + 1
            End If
        End If
        sFolder = Dir()
    Loop
End Sub

