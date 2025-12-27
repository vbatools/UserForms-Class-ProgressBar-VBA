VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProgresBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========================================================================================
' Class: clsProgresBar
' Author: VBATools
' Version: 1.0.2
' Creation Date: 11.12.2025 19:41
' Last Update Date: 11.12.2025 19:41
' License: Apache License
'========================================================================================

Private mFrmForma   As frmProgresBar_2
Private mWidthForm  As Single
Private mTimeStart  As Date
Private mTypeCaptionLabel As enumTypeCaptionLabel
Private mCountItems As Long
Private mLogArr()
Private mStep       As Long

Public Enum enumTypeCaptionLabel
    enNone
    enProcent
    enValue
    enTime
    enProcentValue
    enProcentTime
    enValuTime
    enAll
End Enum

Public Enum enumParametrVersion
    enName = 0
    enAuthor
    enVersion
    enLicense
    enDateOfCreation
    enDateOfUpdate
    enDescription
    enAll
    [_First] = enName
    [_Last] = enAll
End Enum

' Returns version information based on the specified parameter
Public Function Version(ByVal Parametr As enumParametrVersion) As String
    Dim sRes        As String
    Select Case Parametr
        Case enumParametrVersion.enAll:
            Dim i   As Byte
            For i = enumParametrVersion.[_First] To enumParametrVersion.[_Last] - 1
                If sRes <> vbNullString Then sRes = sRes & vbNewLine
                sRes = sRes & VersionItem(i)
            Next i
        Case Else:
            sRes = VersionItem(Parametr)
    End Select
    Version = sRes
End Function

' Helper function to return individual version items
Private Function VersionItem(ByVal Parametr As enumParametrVersion) As String
    Select Case Parametr
        Case enumParametrVersion.enName:
            VersionItem = "Class: clsProgresBar"
        Case enumParametrVersion.enAuthor:
            VersionItem = "Author: VBATools"
        Case enumParametrVersion.enVersion:
            VersionItem = "Version: 1.0.2"
        Case enumParametrVersion.enLicense:
            VersionItem = "License: Apache License"
        Case enumParametrVersion.enDateOfCreation:
            VersionItem = "Date of creation: 11.12.2025 19:41"
        Case enumParametrVersion.enDateOfUpdate:
            VersionItem = "Date of update: 11.12.2025 19:41"
        Case enumParametrVersion.enDescription:
            VersionItem = "Description: Class for displaying a progress bar in VBA. Provides a convenient interface for displaying the progress of long-running operations with the ability to customize the appearance and display accompanying messages."
    End Select
End Function

Property Get Forma() As frmProgresBar_2
    If mFrmForma Is Nothing Then
        Set Forma = New frmProgresBar_2
    Else
        Set Forma = mFrmForma
    End If
End Property

Property Get LogData() As Variant
    LogData = mLogArr
End Property

Property Get TypeCaptionLabel() As enumTypeCaptionLabel
    TypeCaptionLabel = mTypeCaptionLabel
End Property

Public Property Let TypeCaptionLabel(ByRef TypeCaptionLabel As enumTypeCaptionLabel)
    mTypeCaptionLabel = TypeCaptionLabel
End Property

Property Get HeaderCaption() As String
    HeaderCaption = mFrmForma.Caption
End Property

Public Property Let HeaderCaption(ByRef sHeader As String)
    mFrmForma.Caption = sHeader
End Property

Property Get MessageTop() As String
    MessageTop = mFrmForma.MessageTop.Caption
End Property

Public Property Let MessageTop(ByRef sMessage As String)
    mFrmForma.MessageTop.Caption = sMessage
End Property

Property Get MessageBottom() As String
    MessageBottom = mFrmForma.MessageBottom
End Property

Public Property Let MessageBottom(ByRef sMessage As String)
    mFrmForma.MessageBottom.Caption = sMessage
End Property

Property Get PictureShow() As Boolean
    PictureShow = mFrmForma.lbPict.Visible
End Property

Public Property Let PictureShow(ByRef Show As Boolean)
    mFrmForma.lbPict.Visible = Show
End Property

Property Get PictureSimvol() As String
    PictureSimvol = mFrmForma.lbPict.Caption
End Property

Public Property Let PictureSimvol(ByRef PictureSimvol As String)
    mFrmForma.lbPict.Caption = PictureSimvol
End Property

Property Get PictureColor() As XlRgbColor
    PictureColor = mFrmForma.lbPict.ForeColor
End Property

Public Property Let PictureColor(ByRef Color As XlRgbColor)
    mFrmForma.lbPict.ForeColor = Color
End Property

Property Get LineFrontColor() As XlRgbColor
    LineFrontColor = mFrmForma.lbLineFront.BackColor
End Property

Public Property Let LineFrontColor(ByRef Color As XlRgbColor)
    mFrmForma.lbLineFront.BackColor = Color
End Property

Property Get LineBackColor() As XlRgbColor
    LineBackColor = mFrmForma.lbLineBack.BackColor
End Property

Public Property Let LineBackColor(ByRef Color As XlRgbColor)
    mFrmForma.lbLineBack.BackColor = Color
End Property

Property Get TimeWork() As Date
    TimeWork = VBA.Now() - mTimeStart
End Property

Public Sub Initialize(ByRef sHeaderCaption As String, _
        ByRef sMessageTop As String, _
        ByRef sMessageBottom As String, _
        ByRef TypeCaptionLabel As enumTypeCaptionLabel, _
        Optional LineFrontColor As XlRgbColor = -1, _
        Optional LineBackColor As XlRgbColor = -1, _
        Optional sLineFrontSimvol As String = "|", _
        Optional bPictureShow As Boolean = False, _
        Optional sPictureSimvol As String = vbNullString)

    Set mFrmForma = Forma()
    mTypeCaptionLabel = TypeCaptionLabel
    mTimeStart = VBA.Now()
    With mFrmForma
        .Caption = sHeaderCaption
        .MessageTop = sMessageTop
        .MessageBottom = sMessageBottom
        If mCountItems < 1 Then mCountItems = 100
        .lbProcent.Caption = getCaptionLabel(TypeCaptionLabel, 0, mCountItems)
        .lbLineFront.Width = 0
        If sLineFrontSimvol <> vbNullString Then .lbLineFront.Caption = VBA.String$(300, sLineFrontSimvol)
        If LineFrontColor <> -1 Then .lbLineFront.BackColor = LineFrontColor
        If LineBackColor <> -1 Then .lbLineBack.BackColor = LineBackColor
        With .lbPict
            .Caption = sPictureSimvol
            .Left = 0
            .Visible = bPictureShow
        End With
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
        mWidthForm = .Width
        Call .Show(0)
    End With
End Sub

Public Sub Resize(ByVal WidthForm As Single, _
        ByVal HeightMessageTop As Single, _
        ByVal HeightMessageBottom As Single, _
        ByVal HeightLineProgres As Single)

    With mFrmForma
        Dim dbDelta As Single
        dbDelta = WidthForm - .Width
        If WidthForm > mWidthForm Then
            .MessageBottom.Width = dbDelta + .MessageBottom.Width
            .MessageTop.Width = dbDelta + .MessageTop.Width
            .lbLineBack.Width = dbDelta + .lbLineBack.Width
            .lbProcent.Left = .lbLineBack.Left
            .lbProcent.Width = .lbLineBack.Width
            .Width = WidthForm
        End If

        dbDelta = HeightMessageTop - .MessageTop.Height
        If .Height + dbDelta > 0 And .MessageTop.Height + dbDelta > 0 Then
            .Height = .Height + dbDelta
            .MessageTop.Height = .MessageTop.Height + dbDelta

            .MessageBottom.Top = .MessageBottom.Top + dbDelta
            .lbProcent.Top = .lbProcent.Top + dbDelta + 1
            .lbLineFront.Top = .lbLineFront.Top + dbDelta
            .lbLineBack.Top = .lbLineBack.Top + dbDelta
            .lbPict.Top = .lbPict.Top + dbDelta
        End If

        dbDelta = HeightMessageBottom - .MessageBottom.Height
        If .Height + dbDelta > 0 And .MessageBottom.Height + dbDelta > 0 Then
            .Height = .Height + dbDelta
            .MessageBottom.Height = .MessageBottom.Height + dbDelta

            .lbProcent.Top = .lbProcent.Top + dbDelta + 1
            .lbLineFront.Top = .lbLineFront.Top + dbDelta
            .lbLineBack.Top = .lbLineFront.Top
            .lbPict.Top = .lbPict.Top + dbDelta
        End If

        dbDelta = HeightLineProgres - .lbLineFront.Height
        If .Height + dbDelta > 0 And .lbLineFront.Height + dbDelta > 0 Then
            .Height = .Height + dbDelta
            .lbLineFront.Height = .lbLineFront.Height + dbDelta
            .lbLineBack.Height = .lbLineFront.Height
            .lbProcent.Top = .lbLineBack.Top + .lbLineBack.Height / 2 - .lbProcent.Height / 2
        End If

        .lbProcent.Visible = .lbLineFront.Height >= .lbProcent.Height

        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
        .Repaint
    End With
End Sub

Public Function Update(ByVal item As Long, _
        ByVal CountItems As Long, _
        ByVal sMessageBottom As String, _
        Optional LineFrontColor As XlRgbColor = -1, _
        Optional LineBackColor As XlRgbColor = -1) As Boolean

    
    Dim procent As Single
    mCountItems = CountItems
    procent = item / mCountItems
    Select Case procent
        Case Is < 0:
            procent = 0
        Case Is > 1:
            procent = 1
    End Select
    With mFrmForma
        Update = VBA.CBool(.lbClose.Caption)
        .MessageBottom.Caption = sMessageBottom
        .lbLineFront.Width = .lbLineBack.Width * procent
        If .lbPict.Visible Then .lbPict.Left = .lbLineFront.Left + .lbLineFront.Width - 5
        If LineFrontColor <> -1 Then .lbLineFront.BackColor = LineFrontColor
        If LineBackColor <> -1 Then .lbLineBack.BackColor = LineBackColor
        mCountItems = CountItems
        .lbProcent.Caption = getCaptionLabel(mTypeCaptionLabel, procent, mCountItems)

        .lbProcent.ZOrder 0
        .lbPict.ZOrder 0
        .Repaint
        DoEvents

        If mStep = 0 Then ReDim mLogArr(1 To mCountItems, 1 To 3)
        mStep = mStep + 1
        mLogArr(mStep, 1) = sMessageBottom
        mLogArr(mStep, 2) = VBA.Time
        If mStep = 1 Then
            mLogArr(mStep, 3) = Me.TimeWork
        Else
            mLogArr(mStep, 3) = mLogArr(mStep, 2) - mLogArr(mStep - 1, 2)
        End If
    End With
End Function

Private Function getCaptionLabel(ByVal TypeCaptionLabel As enumTypeCaptionLabel, ByVal procent As Single, ByVal CountItems As Long) As String
    Dim sVal        As String
    Select Case TypeCaptionLabel
        Case enumTypeCaptionLabel.enProcent
            sVal = VBA.Format$(procent, "0%")
        Case enumTypeCaptionLabel.enValue
            sVal = VBA.Round(CountItems * procent, 0) & "/" & CountItems
        Case enumTypeCaptionLabel.enTime
            sVal = Me.TimeWork
        Case enumTypeCaptionLabel.enProcentTime
            sVal = VBA.Format$(procent, "0%") & vbTab & Me.TimeWork
        Case enumTypeCaptionLabel.enProcentValue
            sVal = VBA.Format$(procent, "0%") & "  " & VBA.Round(CountItems * procent, 0) & "/" & CountItems
        Case enumTypeCaptionLabel.enValuTime
            sVal = VBA.Round(CountItems * procent, 0) & "/" & CountItems & "  " & Me.TimeWork
        Case enumTypeCaptionLabel.enAll
            sVal = VBA.Format$(procent, "0%") & "  " & VBA.Round(CountItems * procent, 0) & "/" & CountItems & "  " & Me.TimeWork
        Case enumTypeCaptionLabel.enNone
            'nothing
    End Select
    getCaptionLabel = sVal
End Function

Private Sub Class_Terminate()
    mFrmForma.Hide
    Set mFrmForma = Nothing
End Sub
