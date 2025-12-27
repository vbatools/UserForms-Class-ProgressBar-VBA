Attribute VB_Name = "modTest"
Option Explicit

#If Win64 Then
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
#Else
Public Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
#End If

Sub testProgressBar()
    Dim oProg       As clsProgresBar
    Set oProg = New clsProgresBar
    Call oProg.Initialize("text_header", "text msg top", "text msg bottom", _
            enumTypeCaptionLabel.enAll, rgbAzure, rgbWheat, vbNullString, True, "o")

    Call oProg.Resize(500, 50, 50, 22)

    Dim i           As Long
    For i = 1 To 300 Step 5
        If oProg.Update(i, 300, "text msg bottom_" & i) Then
            Set oProg = Nothing
            Exit For
        End If
        Call Sleep(100)
        If i > 150 Then
            oProg.MessageTop = "New size"
            Call oProg.Resize(500, 80, 20, 5)
        End If
    Next i
    Dim arr         As Variant
    If Not oProg Is Nothing Then
        arr = oProg.LogData
        Cells(1, 1).Resize(UBound(arr, 1), UBound(arr, 2)).Value2 = arr
    End If
End Sub
