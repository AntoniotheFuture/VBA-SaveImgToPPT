Attribute VB_Name = "模块1"
Public SavePath As String
Public CoverSave As Boolean
Public SaveImgPath As String
Public ScanTime As Integer
Public ScanFolder As String
Public TempFolder As String
Public ImgName As Integer

Const CTitle = 1
Const CSender = 2
Const CAttm = 3
Const CMinW = 4
Const CMaxW = 5
Const CMinH = 6
Const CMaxH = 7

Public Errors As String


Private Sub GetAtts()
Dim WS As Worksheet
Set WS = ThisWorkbook.Sheets("参数表")
'默认参数
SavePath = ThisWorkbook.Path
CoverSave = False
SaveImgPath = ""
ScanTime = 180
ScanFolder = ""
TempFolder = ThisWorkbook.Path & "\临时"
For x = 2 To WS.UsedRange.Rows.Count
    Select Case WS.Cells(x, 1).Value
    Case "1"
        SavePath = WS.Cells(x, 3).Value
    Case "2"
        CoverSave = WS.Cells(x, 3).Value
    Case "3"
        SaveImgPath = WS.Cells(x, 3).Value
    Case "4"
        Select Case WS.Cells(x, 3).Value
        Case "1小时"
            ScanTime = 60
        Case "3小时"
            ScanTime = 180
        Case "6小时"
            ScanTime = 360
        Case "1天"
            ScanTime = 1440
        End Select
    Case "5"
        ScanFolder = WS.Cells(x, 3).Value
    Case "6"
        TempFolder = WS.Cells(x, 3).Value
    End Select
Next
End Sub

Sub GetMails()
Dim PPTApp As PowerPoint.Application
Dim PPTFile As PowerPoint.Presentation
Dim olMail As Outlook.MailItem
Dim OLF As Outlook.MAPIFolder
Dim Emails
Dim WS As Worksheet
Dim TWB As Workbook
Dim TWS As Worksheet
Dim Attm As Outlook.Attachment
Dim DoLoad As Boolean
Dim Title As String
Dim Sender As String
Dim Attms As String
Dim AttmArr
Dim PMinW As Integer
Dim PMaxW As Integer
Dim PMinH As Integer
Dim PMaxH As Integer
Dim MailC As Collection
Dim Hit As Boolean
Dim Sh As Shape
Dim SC As Integer
Dim TotalRow As Integer
Errors = ""
'On Error Resume Next

GetAtts
'判断邮箱状态
SC = 0
If ScanFolder = "" Then
    Set OLF = GetObject("", "Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
Else
    'Set a = GetObject("", "Outlook.Application").GetNamespace("MAPI").Folders(1).Folders
    
    PArr = Split(ScanFolder, ">")
    Set OLF = GetObject("", "Outlook.Application").GetNamespace("MAPI").Folders(PArr(0))
    If UBound(PArr) > 0 Then
        For i = 1 To UBound(PArr)
             Set OLF = OLF.Folders(PArr(i))
        Next
    End If
End If
If TypeName(OLF) = "Nothing" Then
    Errors = Errors & "无法获得邮件，请检查权限和邮箱是否已经打开" & Chr(10)
    GoTo skip1
End If
Emails = OLF.Items.Count
If Emails = 0 Then
    Errors = Errors & "无法获得邮件，请检查收件箱是否有邮件" & Chr(10)
    GoTo skip1
End If

UserForm1.Show (0)
'将邮件存入变量
Set MailC = New Collection
For i = 1 To Emails
    ShowStatus "扫描邮件中", i / Emails, i & "/" & Emails
    a = DateAdd("n", ScanTime * -1, Now())
    b = OLF.Items(i).ReceivedTime
    If OLF.Items(i).ReceivedTime > a Then
        MailC.Add OLF.Items(i)
    Else
        'Exit For
    End If
Next
If Dir(SavePath, vbDirectory) = "" Then
    Errors = Errors & "存储路径无效" & Chr(10)
    GoTo skip1
End If

If Dir(TempFolder, vbDirectory) = "" Then
    Errors = Errors & "临时目录不存在，请先创建" & Chr(10)
    GoTo skip1
End If
Set WS = ThisWorkbook.Sheets("规则表")

If SaveImgPath <> "" Then
    If Dir(SaveImgPath, vbDirectory) = "" Then
        Errors = Errors & "找不到图片保存路径" & Chr(10)
        GoTo skip1
    End If
End If


Set PPTApp = New PowerPoint.Application
Set PPTFile = PPTApp.Presentations.Add
TotalRow = WS.UsedRange.Rows.Count

ImgName = 1
For x = 2 To TotalRow
    ShowStatus "处理邮件中", x - 1 / TotalRow - 1, x - 1 & "/" & TotalRow - 1
    DoLoad = True
    If WS.Cells(x, CTitle).Value = "" Then
        Errors = Errors & "顺序" & x - 1 & "邮件主题不能为空" & Chr(10)
        DoLoad = False
    End If
    If DoLoad Then
        Title = WS.Cells(x, CTitle).Value
        Sender = ""
        Attms = ""
        PMinW = 0
        PMaxW = 0
        PMinH = 0
        PMaxH = 0
        If WS.Cells(x, CSender).Value <> "" Then Sender = WS.Cells(x, CSender).Value
        If WS.Cells(x, CAttm).Value <> "" Then Attms = WS.Cells(x, CAttm).Value
        If WS.Cells(x, CMinW).Value <> "" Then PMinW = CInt(WS.Cells(x, CMinW).Value)
        If WS.Cells(x, CMaxW).Value <> "" Then PMaxW = CInt(WS.Cells(x, CMaxW).Value)
        If WS.Cells(x, CMinH).Value <> "" Then PMinH = CInt(WS.Cells(x, CMinH).Value)
        If WS.Cells(x, CMaxH).Value <> "" Then PMaxH = CInt(WS.Cells(x, CMaxH).Value)
    End If
    
    
    '遍历邮件
    For Each olMail In MailC
        Hit = True
        If Not olMail.Subject Like Title Then Hit = False
            
        'If Title <> olMail.Subject Then Hit = False
        If Sender <> "" Then
            Hit = False
            If Sender = olMail.Sender.Address Then Hit = True
        End If
        If Hit Then
            If olMail.Attachments.Count = 0 Then
                Hit = False
                Errors = Errors & "顺序" & x - 1 & "邮件没有附件" & Chr(10)
             End If
        End If
        If Hit Then
             If Attms <> "" Then
                AttmArr = Split(Attms, ";")
                For Each Attm In olMail.Attachments
                    For Each a In AttmArr
                        If a = Attm.DisplayName Then
                            C = TempFolder & "\" & Year(Date) & Month(Date) & Day(Date) & Hour(Now()) & Minute(Now()) & Second(Now()) & Attm.Filename
                            Attm.SaveAsFile C
                            Set TWB = Workbooks.Open(C)
                            GetAllImgIntoPPT TWB, PPTFile, PMinW, PMaxW, PMinH, PMaxH, olMail.Subject
                            TWB.Close False
                        End If
                    Next
                    SC = SC + 1
                Next
            Else
                For Each Attm In olMail.Attachments
                    C = TempFolder & "\" & Year(Date) & Month(Date) & Day(Date) & Hour(Now()) & Minute(Now()) & Second(Now()) & Attm.Filename
                    Attm.SaveAsFile C
                    Set TWB = Workbooks.Open(C)
                    GetAllImgIntoPPT TWB, PPTFile, PMinW, PMaxW, PMinH, PMaxH, olMail.Subject
                    TWB.Close False
                Next
                SC = SC + 1
            End If
        End If
    Next
Next
If SC > 0 Then
    'PPTFile.SaveAs SavePath & "\" & Year(Date) & Month(Date) & Day(Date) & "合并邮件.pptx"
    Set PPTFile = Nothing
    Set PPTApp = Nothing
Else
    PPTFile.Close
    Set PPTFile = Nothing
    Set PPTApp = Nothing
End If

skip1:
If Errors <> "" Then MsgBox (Errors)
If SC > 0 Then
    MsgBox ("成功处理" & SC & "封邮件")
Else
    MsgBox ("未处理任何邮件")
End If
End Sub

Private Sub ShowStatus(Title As String, Progress As Double, Info As String)
With UserForm1
    .Caption = Title
    .Label1.Width = Progress * .Width
    .Label2.Caption = Info
    DoEvents
End With
End Sub

Private Sub GetAllImgIntoPPT(WB As Workbook, PPT As PowerPoint.Presentation, MinW As Integer, MaxW As Integer, MinH As Integer, MaxH As Integer, Optional Title As String)
Dim WS As Worksheet
Dim Sh As Shape
Dim Hit As Boolean
Dim PPTS As PowerPoint.Slide
Dim TB As PowerPoint.Shape
Dim pptLayout As CustomLayout
Dim i As Integer
Dim S As PowerPoint.Slide
'PPT.Slides.AddSlide 0.7
'a = ppLayoutCustom
'Set pptLayout = ppLayoutCustom
'Set pptLayout = PPT.Slides(0).CustomLayout
For Each WS In WB.Sheets
    For Each Sh In WS.Shapes
        i = 0
        If Sh.Type = msoPicture Then
            i = i + 1
            Hit = True
            If MinW <> 0 Then
                If Sh.Width < MinW Then Hit = False
            End If
            If MaxW <> 0 Then
                If Sh.Width > MaxW Then Hit = False
            End If
            If MinH <> 0 Then
                If Sh.Height < MinH Then Hit = False
            End If
            If MaxH <> 0 Then
                If Sh.Height > MaxH Then Hit = False
            End If
            If Hit Then
                Set PPTS = PPT.Slides.Add(PPT.Slides.Count + 1, ppLayoutCustom)
                'Set PPTS = PPT.Slides.AddSlide(PPT.Slides.Count + 1, pptLayout)
                Set TB = PPTS.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, PPTS.Master.Width, 50)
                TB.TextFrame2.TextRange.Text = Title & ">" & WB.Name & ">" & WS.Name & ">" & i
                TB.Left = 0
                TB.Top = 0
                Sh.Copy
                PPTS.Shapes.Paste
                'If SaveImgPath <> "" Then
                '    PPTS.Shapes(PPTS.Shapes.Count).Export SaveImgPath & "\" & ImgName & ".png", ppShapeFormatPNG
                '    ImgName = ImgName + 1
                'End If
            End If
        End If
    Next
Next

If SaveImgPath <> "" Then
    For Each S In PPT.Slides
        Set TB = S.Shapes(3)
        TB.Export SaveImgPath & "\" & ImgName & ".png", ppShapeFormatPNG
        ImgName = ImgName + 1
    Next
End If

End Sub

