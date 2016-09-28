VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Install"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Const WorkbookPathPlaceholder As String = ""

Public Enum Color
    Green
    Red
End Enum

Public Property Get Version() As String
    Version = Me.Range("Version").Value
End Property
Public Property Let Version(Value As String)
    Me.Range("Version").Value = Value
End Property

Public Property Get WorkbookPath() As String
    WorkbookPath = Me.Range("WorkbookPath").Value
    If WorkbookPath = WorkbookPathPlaceholder Then: WorkbookPath = ""
End Property
Public Property Let WorkbookPath(Value As String)
    Me.Range("WorkbookPath").Value = Value
End Property

Public Property Get Selections() As VBAWebSelections
    Selections.Src = Me.Range("InstallSrc").Value
    Selections.VBADictionary = Me.Range("InstallVBADictionary").Value
    Selections.AsyncWrapper = Me.Range("InstallAsyncWrapper").Value
    Selections.HttpBasicAuthenticator = Me.Range("InstallHttpBasic").Value
    Selections.OAuth1Authenticator = Me.Range("InstallOAuth1").Value
    Selections.OAuth2Authenticator = Me.Range("InstallOAuth2").Value
    Selections.DigestAuthenticator = Me.Range("InstallDigest").Value
    Selections.WindowsAuthenticator = Me.Range("InstallWindows").Value
    Selections.GoogleAuthenticator = Me.Range("InstallGoogle").Value
    Selections.FacebookAuthenticator = Me.Range("InstallFacebook").Value
    Selections.TwitterAuthenticator = Me.Range("InstallTwitter").Value
    Selections.TodoistAuthenticator = Me.Range("InstallTodoist").Value
    Selections.EmptyAuthenticator = Me.Range("InstallEmpty").Value
End Property
Public Property Let Selections(Values As VBAWebSelections)
    Me.Range("InstallSrc").Value = Values.Src
    Me.Range("InstallVBADictionary").Value = Values.VBADictionary
    Me.Range("InstallAsyncWrapper").Value = Values.AsyncWrapper
    Me.Range("InstallHttpBasic").Value = Values.HttpBasicAuthenticator
    Me.Range("InstallOAuth1").Value = Values.OAuth1Authenticator
    Me.Range("InstallOAuth2").Value = Values.OAuth2Authenticator
    Me.Range("InstallDigest").Value = Values.DigestAuthenticator
    Me.Range("InstallWindows").Value = Values.WindowsAuthenticator
    Me.Range("InstallGoogle").Value = Values.GoogleAuthenticator
    Me.Range("InstallFacebook").Value = Values.FacebookAuthenticator
    Me.Range("InstallTwitter").Value = Values.TwitterAuthenticator
    Me.Range("InstallTodoist").Value = Values.TodoistAuthenticator
    Me.Range("InstallEmpty").Value = Values.EmptyAuthenticator
End Property

Public Property Let ProgressBarWidth(Value As Double)
    Dim ProgressBar As Shape
    Dim MaxWidth As Double
    
    Set ProgressBar = Me.Shapes("ProgressBar")
    MaxWidth = Me.Shapes("ProgressBarContainer").Width - 2
    
    ' 100%: Width = 255, 0%: Width = 0
    If Value > 1 Then
        ProgressBar.Width = MaxWidth
    ElseIf Value < 0 Then
        ProgressBar.Width = 0
    Else
        ProgressBar.Width = Value * MaxWidth
    End If
End Property

Public Property Let ProgressBarColor(Value As Color)
    Dim ColorCode As Long

    Select Case Value
    Case Color.Green
        ColorCode = RGB(0, 113, 51)
    Case Color.Red
        ColorCode = RGB(204, 0, 0)
    End Select
    
    Me.Shapes("ProgressBar").Fill.ForeColor.RGB = ColorCode
    Me.Shapes("ProgressBarContainer").Line.ForeColor.RGB = ColorCode
End Property

Public Property Let ProgressDescription(Value As String)
    Me.Range("ProgressDescription").Value = Value
End Property

Public Property Let ProgressDetails(Value As String)
    Me.Range("ProgressDetails").Value = Value
End Property

Public Sub Browse()
    Dim Result As String
    Result = BrowseForProjectPath
    
    If Result <> "" Then
        Me.WorkbookPath = Result
        Me.ResetProgressBar
    End If
End Sub

Public Sub Install()
    On Error GoTo ErrorHandling
    
    Me.ResetProgressBar
    VBAWebInstaller.InstallSelections Me.WorkbookPath, Me.Selections
    Exit Sub
    
ErrorHandling:

    Me.ProgressBarWidth = 100#
    Me.ProgressDescription = "ERROR"
    Me.ProgressDetails = "ERROR " & Err.Number & ": " & Err.Description
    Me.ProgressBarColor = Color.Red
End Sub

Public Sub ShowProgress(TotalCount As Long, CompleteCount As Long)
    If TotalCount > 0 Then
        Me.ProgressBarWidth = CompleteCount / TotalCount
    End If
    
    If CompleteCount = TotalCount Then
        Me.ProgressDescription = "Done!"
    Else
        Me.ProgressDescription = "Installing... (" & CompleteCount & " of " & TotalCount & ")"
    End If
End Sub

Public Sub Reset()
    Dim DefaultSelections As VBAWebSelections

    Me.WorkbookPath = WorkbookPathPlaceholder
    
    DefaultSelections.Src = True
    DefaultSelections.VBADictionary = True
    DefaultSelections.AsyncWrapper = False
    DefaultSelections.OAuth1Authenticator = False
    DefaultSelections.OAuth2Authenticator = False
    DefaultSelections.DigestAuthenticator = False
    DefaultSelections.GoogleAuthenticator = False
    DefaultSelections.FacebookAuthenticator = False
    DefaultSelections.TwitterAuthenticator = False
    DefaultSelections.EmptyAuthenticator = False
    
    Me.Selections = DefaultSelections
    
    Me.ResetProgressBar
End Sub

Public Sub ResetProgressBar()
    Me.ProgressBarWidth = 0#
    Me.ProgressDescription = ""
    Me.ProgressDetails = ""
    Me.ProgressBarColor = Color.Green
End Sub

Public Function BrowseForProjectPath() As String
#If Mac Then
    Dim Script As String
    Dim Result As String
    
    ' TODO Add Access install support on Mac
    Script = _
        "set applescript's text item delimiters to "",""" & vbNewLine & _
        "set filepath to (choose file of type " & _
        "{""com.microsoft.Excel.xls"",""org.openxmlformats.spreadsheetml.sheet.macroenabled""} " & _
        "with prompt ""Choose Project (Excel or Access)"" multiple selections allowed false) as string" & vbNewLine & _
        "set applescript's text item delimiters to """"" & vbNewLine & _
        "return filepath"
    
    On Error Resume Next
    Result = MacScript(Script)
    
    If Result <> "" Then
        BrowseForProjectPath = Result
    End If
#Else
    Dim Result As String
    Result = Application.GetOpenFilename( _
        FileFilter:="Excel or Access Files, *.xl*;*.acc*", _
        Title:="Choose Project (Excel or Access)", _
        MultiSelect:=False)
    
    If Result <> "" And Result <> "False" Then
        BrowseForProjectPath = Result
    End If
#End If
End Function
