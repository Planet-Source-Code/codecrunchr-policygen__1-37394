VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Policy Number Generator"
   ClientHeight    =   3510
   ClientLeft      =   435
   ClientTop       =   2055
   ClientWidth     =   6270
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   6270
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   2880
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reprints"
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process"
      Enabled         =   0   'False
      Height          =   615
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   615
      Index           =   0
      Left            =   4440
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Policy Build Information"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6015
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3600
         TabIndex        =   5
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":030A
         Left            =   3600
         List            =   "Form1.frx":0320
         TabIndex        =   1
         Text            =   "NONE"
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Number of Policies to generate:"
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   1080
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Select Business Area:"
         Height          =   195
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   1545
      End
   End
   Begin MSComDlg.CommonDialog PrinterControl 
      Left            =   3960
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Policy Number Generator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private adoPrimaryRS        As New ADODB.Recordset
Private WordDoc             As Word.Application
Private thisDoc             As Document
Dim blnWordDoc              As Boolean
Private prnTime             As Date
Private strPolicyNumber     As String
Private IsReprint           As Boolean
Private Reply

Enum TemplateTypes
    NBintlPatchPage = 0
    NBLifePatchPage = 1
    NBAnnPatchPage = 2
    CSintlPatchPage = 3
    CSLifePatchPage = 4
    CSAnnPatchPage = 5
End Enum

Public Sub CopyTemplate()
    ' Copy the template for duplication.
    With WordDoc.Selection
        .WholeStory
        .Copy
    End With
End Sub

Public Sub DuplicatePage()
    ' Add page and paste a copy of the template on it.
    With WordDoc.Selection
        .MoveRight Unit:=wdCharacter, Count:=1
        .InsertBreak Type:=wdPageBreak
        .Paste
    End With
End Sub

Public Sub MakePolicyNumbers()
    ' Scan the paragraphs and replace "0000000000" with the new policy number.
    Dim i As Integer
    Dim x As String
    x = strPolicyNumber
    With WordDoc.ActiveDocument
        For i = 1 To .Paragraphs.Count
            .Paragraphs(i).Range.Select
            If Trim(Left(WordDoc.Selection.Text, 10)) = "0000000000" Then
                If Not IsReprint Then
                     x = "0" & Val(x) + 1
                Else
                    x = "0" & Val(x)
                End If
                WordDoc.Selection.TypeText Text:=x
            End If
        Next
    End With
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0 ' Exit
            SaveSettings
            End
        Case 1 ' Process
            If Val(Text1.Text) > 0 Then
                Me.MousePointer = 11
                IsReprint = False
                BuildSelection
            Else
                MsgBox "Please Enter number of Policy Numbers!", vbCritical, "Policy Number Generator"
            End If
        Case 2 ' Reprints
            Me.MousePointer = 11
            IsReprint = True
            BuildSelection
            IsReprint = False
    End Select
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, "Settings", "Top", Top
    SaveSetting App.EXEName, "Settings", "Left", Left
End Sub
Private Sub BuildSelection()
    Dim i As Integer
    If IsReprint Then
        i = 1
    Else
        i = Text1.Text
    End If
    Select Case Combo1.Text
        Case "NONE"
            MsgBox "Please select a Business Area!", vbCritical, "Policy Number Generator"
            Me.MousePointer = 0
            Exit Sub
        Case "NB-Annuity"
            MakePatchPages NBAnnPatchPage, i
        Case "NB-Intl"
            MakePatchPages NBintlPatchPage, i
        Case "NB-Life"
            MakePatchPages NBLifePatchPage, i
        Case "CS-Annuity"
            MakePatchPages CSAnnPatchPage, i
        Case "CS-Life"
            MakePatchPages CSLifePatchPage, i
        Case "CS-Intl"
            MakePatchPages CSintlPatchPage, i
    End Select
End Sub

Private Sub MakePatchPages(BusinessType As TemplateTypes, Pages As Integer)
    ' Pass errors
    On Error Resume Next
    Dim i As Index
    Dim NewPolNum As String
    Set WordDoc = GetObject("Word.Application")
    If WordDoc Is Nothing Then
        Set WordDoc = CreateObject("Word.Application")
        If WordDoc Is Nothing Then
            MsgBox "Could not start Word. Make sure application is available.", , "Policy Number Generator"
            Exit Sub
        End If
    End If
    WordDoc.Visible = False
    blnWordDoc = True
    Select Case BusinessType
        Case 0      ' NBintlPatchPage
            Set thisDoc = WordDoc.Documents.Add(App.Path & "\NBintlPatchPage.dot")
        Case 1      ' NBLifePatchPage
            Set thisDoc = WordDoc.Documents.Add(App.Path & "\NBLifePatchPage.dot")
        Case 2      ' NBAnnPatchPage
            Set thisDoc = WordDoc.Documents.Add(App.Path & "\NBAnnPatchPage.dot")
        Case 3      ' CSintlPatchPage
            Set thisDoc = WordDoc.Documents.Add(App.Path & "\CSintlPatchPage.dot")
        Case 4      ' CSLifePatchPage
            Set thisDoc = WordDoc.Documents.Add(App.Path & "\CSLifePatchPage.dot")
        Case 5      ' CSAnnPatchPage
            Set thisDoc = WordDoc.Documents.Add(App.Path & "\CSAnnPatchPage.dot")
    End Select
    CopyTemplate
    Do While Pages > 1
        DuplicatePage
        Pages = Pages - 1
    Loop
    If Not IsReprint Then
'        adoPrimaryRS.Requery
        strPolicyNumber = GetSetting(App.EXEName, "Settings", "PolicyNumber", "0100000000") 'adoPrimaryRS.Fields("PolicyNumber")
        NewPolNum = Val(strPolicyNumber) + Val(Text1.Text)
        ' incriment the policy number.
        ' If you store the policy number in the registry.
        If Len(NewPolNum) = 10 Then
            SaveSetting App.EXEName, "Settings", "PolicyNumber", NewPolNum
        Else
            SaveSetting App.EXEName, "Settings", "PolicyNumber", "0" & NewPolNum
        End If
        ' if you store the policy number in a database.
'        With adoPrimaryRS
'            If Len(NewPolNum) = 10 Then
'                .Fields("PolicyNumber").Value = NewPolNum
'            Else
'                .Fields("PolicyNumber").Value = "0" & NewPolNum
'            End If
'            .Update
'        End With
    Else
        strPolicyNumber = InputBox("Enter Policy Number for Reprint:", "Reprints")
    End If
    MakePolicyNumbers
    PrinterControl.ShowPrinter
    thisDoc.PrintOut True, True
    prnTime = Time
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    ' If you use a database to store the policynumber.
'    Dim db As Connection
'    Set db = New Connection
'    db.CursorLocation = adUseClient
'    db.Open "PROVIDER=MSDASQL;dsn=PolicyGen;uid=UserID;pwd=;database=PolicyGen;"
'    Set adoPrimaryRS = New Recordset
'    adoPrimaryRS.Open "select PolicyNumber from PolicyStatus", db, adOpenStatic, adLockOptimistic
'    strPolicyNumberadoPrimaryRS.Fields("PolicyNumber") = adoPrimaryRS.Fields("PolicyNumber")

    ' If you use the registry to store the poicynumber.
    strPolicyNumber = GetSetting(App.EXEName, "Settings", "PolicyNumber", "0100000000")
    Left = GetSetting(App.EXEName, "Settings", "Left", "0")
    Top = GetSetting(App.EXEName, "Settings", "Top", "0")
End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    ' If you use a database to store the policynumber.
'    Set adoPrimaryRS = Nothing
    SaveSettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ' If you use a database to store the policynumber.
'    Set adoPrimaryRS = Nothing
    SaveSettings
End Sub

Private Sub Text1_Change()
    ' Number of prints must be larger than none.
    If Text1.Text > "0" Then _
    Command1(1).Enabled = True Else _
    Command1(1).Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    ' Only allow numeric keys.
    If Not KeyAscii = 8 Then
        If KeyAscii < 48 Or KeyAscii > 57 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    ' Wait for print status to be finished.
    If Not WordDoc.BackgroundPrintingStatus <> 0 Then
        thisDoc.Close False
        WordDoc.Quit
        Set WordDoc = Nothing
        blnWordDoc = False
        MsgBox "Job Completed Successfully", , "Policy Number Generator"
        Me.MousePointer = 0
        Timer1.Enabled = False
    Else
        If Minute(Time - prnTime) > 1 Then
            Reply = MsgBox("Word is taking too long to print." & vbCrLf & "Do you want to quit?", vbYesNo, "Policy Number Generator")
            If Reply = vbYes Then
                thisDoc.Close False
                WordDoc.Quit
                Set WordDoc = Nothing
                blnWordDoc = False
                MsgBox "Job was lost, Please try again later.", vbCritical, "Policy Number Generator"
                Me.MousePointer = 0
                Timer1.Enabled = False
            End If
        End If
    End If
End Sub
