# VBA
Option Explicit

Sub test()
Dim i%, j%, rng As Range, lrow%
lrow = Worksheets("Sheet1").Range("A65536").End(xlUp).Row

'1 info source
    If Range("E2") <> Range("E" & lrow) Then
        If InStr(Range("E2"), "B1") > 0 Or InStr(Range("E2"), "C1") > 0 Then
            If InStr(Range("E" & lrow), "A1") > 0 And InStr(Range("?"), "This case turned out to be spontaneous case.") > 0 Then
                Range("?").Interior.ColorIndex = 6
                With Range("?").AddComment
                    .Visible = False
                    .Text "This case turned out to be spontaneous case."
                End With
            End If
        ElseIf InStr(Range("E" & lrow), "A1") > 0 Then
            If (InStr(Range("E2"), "B1") > 0 Or InStr(Range("E2"), "C1") > 0) And InStr(Range("?"), "This case turned out to be solicited case.") > 0 Then
                Range("?").Interior.ColorIndex = 6
                With Range("?").AddComment
                    .Visible = False
                    .Text "This case turned out to be solicited case."
                End With
            End If
        End If
    End If

'2 age group
    If Range("G" & lrow) <> "" Then
        If Range("F" & lrow) <> "" Then
            Range("F" & lrow).Interior.ColorIndex = 6
            With Range("F" & lrow).AddComment
                .Visible = False
                .Text "Delete"
            End With
        End If
    End If

'3
'    If Range("J" & lrow) Is Not Nothing Then
'        If Range("S" & lrow) Then
'        End If
'    End If

'4
    For i = lrow To Worksheets("Sheet1").Range("AE65536").End(xlUp).Row
        If Range("AJ" & i) = "" Then
            If Range("AO" & i) <> "不明" Then
                Range("AO" & i).Interior.ColorIndex = 6
                With Range("AO" & i).AddComment
                    .Visible = False
                    .Text "不明"
                End With
            End If
            If Range("AP" & i) <> "Unknown" Then
                Range("AP" & i).Interior.ColorIndex = 6
                With Range("AP" & i).AddComment
                    .Visible = False
                    .Text "Unknown"
                End With
            End If
        End If
    Next













End Sub
