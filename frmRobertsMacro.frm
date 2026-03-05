VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRobertsMacro
   Caption         =   "Robert's Definitions Tool"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   6240
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblDefsCol
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   252
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   508
      Caption         =   "Definitions column (e.g. AL):"
   End
   Begin MSForms.TextBox txtDefinitionsCol
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   3360
      _ExtentX        =   5930
      _ExtentY        =   556
   End
   Begin MSForms.Label lblClauseCols
      Height          =   288
      Left            =   240
      TabIndex        =   2
      Top             =   732
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   508
      Caption         =   "Clause columns (e.g. H:AK):"
   End
   Begin MSForms.TextBox txtClauseCols
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   3360
      _ExtentX        =   5930
      _ExtentY        =   556
   End
   Begin MSForms.Label lblRowRanges
      Height          =   288
      Left            =   240
      TabIndex        =   4
      Top             =   1212
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   508
      Caption         =   "Rows to process (e.g. 2:500):"
   End
   Begin MSForms.TextBox txtRowRanges
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Top             =   1200
      Width           =   3360
      _ExtentX        =   5930
      _ExtentY        =   556
   End
   Begin MSForms.CommandButton btnSaveChoices
      Caption         =   "Save choices"
      Height          =   420
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   741
   End
   Begin MSForms.CommandButton btnRunRobertsMacro
      Caption         =   "Run RobertsMacro"
      Default         =   -1
      Height          =   420
      Left            =   1800
      TabIndex        =   7
      Top             =   1920
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   741
   End
   Begin MSForms.CommandButton btnReportBug
      Caption         =   "Report a bug"
      Height          =   420
      Left            =   3840
      TabIndex        =   8
      Top             =   1920
      Width           =   1320
      _ExtentX        =   2329
      _ExtentY        =   741
   End
   Begin MSForms.CommandButton btnClose
      Cancel          =   -1
      Caption         =   "Close"
      Height          =   420
      Left            =   5280
      TabIndex        =   9
      Top             =   1920
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   741
   End
End
Attribute VB_Name = "frmRobertsMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    '=== Resize the form ===
    Me.Width  = 493.5
    Me.Height = 447

    '=== Load saved preferences ===
    Dim defsCol As String, clauseCols As String, rowRanges As String
    LoadUserChoices defsCol, clauseCols, rowRanges

    '--- Shared colour values ---
    Dim cWhite  As Long : cWhite  = RGB(255, 255, 255)
    Dim cBlack  As Long : cBlack  = RGB(0, 0, 0)
    Dim cYellow As Long : cYellow = RGB(255, 255, 153)  ' button background
    Dim cNavy   As Long : cNavy   = RGB(0, 0, 128)      ' button text

    '=== Style existing input labels ===
    With Me.lblDefsCol
        .Caption   = "DEFINITIONS" & vbLf & "COLUMN"
        .Left = 210 : .Top = 110 : .Width = 145 : .Height = 52
        .TextAlign = 2   ' fmTextAlignCenter
        .WordWrap  = True
        .BackStyle = 0   ' fmBackStyleTransparent
        .ForeColor = cWhite
    End With
    ApplyFont Me.lblDefsCol, "Engravers MT", 14, True, False, "Georgia"

    With Me.lblClauseCols
        .Caption   = "CLAUSE" & vbLf & "COLUMNS"
        .Left = 210 : .Top = 183 : .Width = 145 : .Height = 52
        .TextAlign = 2
        .WordWrap  = True
        .BackStyle = 0
        .ForeColor = cWhite
    End With
    ApplyFont Me.lblClauseCols, "Engravers MT", 14, True, False, "Georgia"

    With Me.lblRowRanges
        .Caption   = "ROW RANGES"
        .Left = 210 : .Top = 258 : .Width = 145 : .Height = 30
        .TextAlign = 2
        .BackStyle = 0
        .ForeColor = cWhite
    End With
    ApplyFont Me.lblRowRanges, "Engravers MT", 14, True, False, "Georgia"

    '=== Style existing textboxes ===
    With Me.txtDefinitionsCol
        .Left = 362 : .Top = 113 : .Width = 112 : .Height = 28
        .BackColor     = cBlack
        .ForeColor     = cWhite
        .BorderStyle   = 1   ' fmBorderStyleSingle
        .BorderColor   = cWhite
        .SpecialEffect = 0   ' fmSpecialEffectFlat
    End With
    If defsCol <> "" Then Me.txtDefinitionsCol.Text = defsCol

    With Me.txtClauseCols
        .Left = 362 : .Top = 186 : .Width = 112 : .Height = 28
        .BackColor     = cBlack
        .ForeColor     = cWhite
        .BorderStyle   = 1
        .BorderColor   = cWhite
        .SpecialEffect = 0
    End With
    If clauseCols <> "" Then Me.txtClauseCols.Text = clauseCols

    With Me.txtRowRanges
        .Left = 362 : .Top = 258 : .Width = 112 : .Height = 28
        .BackColor     = cBlack
        .ForeColor     = cWhite
        .BorderStyle   = 1
        .BorderColor   = cWhite
        .SpecialEffect = 0
    End With
    If rowRanges <> "" Then Me.txtRowRanges.Text = rowRanges

    '=== Style existing buttons ===
    Me.btnReportBug.Visible = False   ' not shown in this design

    With Me.btnRunRobertsMacro
        .Caption       = "Run Macro"
        .Left = 122 : .Top = 350 : .Width = 120 : .Height = 48
        .BackColor     = cYellow
        .ForeColor     = cNavy
        .SpecialEffect = 0
    End With
    ApplyFont Me.btnRunRobertsMacro, "Arial Rounded MT Bold", 24, True, True, ""

    With Me.btnSaveChoices
        .Caption       = "Save Choices"
        .Left = 247 : .Top = 350 : .Width = 120 : .Height = 48
        .BackColor     = cYellow
        .ForeColor     = cNavy
        .SpecialEffect = 0
    End With
    ApplyFont Me.btnSaveChoices, "@HGSoeiKakupoptai", 12, True, False, "Trebuchet MS"

    With Me.btnClose
        .Caption       = "Close"
        .Left = 372 : .Top = 350 : .Width = 104 : .Height = 48
        .BackColor     = cYellow
        .ForeColor     = cNavy
        .SpecialEffect = 0
    End With
    ApplyFont Me.btnClose, "@HGSoeiKakupoptai", 16, True, True, "Trebuchet MS"

    '=== Create title label ("Definition extractor") ===
    Dim lblTitle As MSForms.Label
    Set lblTitle = Me.Controls.Add("Forms.Label.1", "lblTitle", True)
    With lblTitle
        .Caption   = "Definition" & vbLf & "extractor"
        .Left = 12 : .Top = 10 : .Width = 175 : .Height = 75
        .WordWrap  = True
        .BackStyle = 0
        .ForeColor = cWhite
    End With
    ApplyFont lblTitle, "Engravers MT", 20, True, True, "Georgia"

    '=== Create paragraph description label ===
    Dim lblDesc As MSForms.Label
    Set lblDesc = Me.Controls.Add("Forms.Label.1", "lblDescription", True)
    With lblDesc
        .Caption  = "This macro automates the scraping and formatting of " & _
                    "definitions within prepopulated cells. You must first " & _
                    "insert all your definitions in a column, and each " & _
                    "definition must take the format of [Defined Term] means " & _
                    "[definition]. Separate each definition with a pipe " & _
                    "character (""|""). The macro will skip rows where " & _
                    "definitions have already been entered. See the one-pager, " & _
                    "or the README file on GitHub, for more information."
        .Left = 12 : .Top = 92 : .Width = 185 : .Height = 240
        .WordWrap  = True
        .BackStyle = 0
        .ForeColor = cWhite
    End With
    ApplyFont lblDesc, "Arial Rounded MT", 12, True, False, "Arial"

    '=== Create rm_macro logo image (clip -- no stretch -- white background) ===
    Dim imgLogo As MSForms.Image
    Set imgLogo = Me.Controls.Add("Forms.Image.1", "imgRmMacro", True)
    With imgLogo
        .Left            = 10 : .Top = 350 : .Width = 105 : .Height = 48
        .PictureSizeMode = 0   ' fmPictureSizeModeClip -- actual size, no stretch
        .BackColor       = RGB(255, 255, 255)
        .BorderStyle     = 0   ' fmBorderStyleNone
    End With
    Dim logoPath As String
    logoPath = ThisWorkbook.Path & Application.PathSeparator & _
               "images" & Application.PathSeparator & "rm_macro.jpg"
    If Dir(logoPath) <> "" Then
        imgLogo.Picture = LoadPicture(logoPath)
    End If

    '=== Create Background_image and send to back ===
    Dim bg As MSForms.Image
    Set bg = Me.Controls.Add("Forms.Image.1", "Background_image", True)
    With bg
        .Left            = 0
        .Top             = 0
        .Width           = Me.InsideWidth
        .Height          = Me.InsideHeight
        .PictureSizeMode = 1   ' fmPictureSizeModeStretch
        .ZOrder 1              ' send to back
    End With
    Dim imgPath As String
    imgPath = ThisWorkbook.Path & Application.PathSeparator & _
              "images" & Application.PathSeparator & "background.jpg"
    If Dir(imgPath) <> "" Then
        bg.Picture = LoadPicture(imgPath)
    End If
End Sub

' Apply a font to a control; if the preferred font is missing and a fallback
' is supplied, the fallback is used instead.
Private Sub ApplyFont(ByVal ctrl As Object, _
                      ByVal fontName As String, _
                      ByVal fontSize As Single, _
                      ByVal isBold As Boolean, _
                      ByVal isItalic As Boolean, _
                      ByVal fallbackFont As String)
    Dim chosen As String
    chosen = fontName
    If fallbackFont <> "" Then
        If Not FontExists(fontName) Then chosen = fallbackFont
    End If
    With ctrl.Font
        .Name   = chosen
        .Size   = fontSize
        .Bold   = isBold
        .Italic = isItalic
    End With
End Sub

' Returns True if the named font is installed on this machine.
' Technique: temporarily create a hidden label, apply the font, and check
' whether Windows accepted the name (it substitutes a different font if not).
Private Function FontExists(ByVal fontName As String) As Boolean
    Const TMP_NAME As String = "_tmpFontCheck_"
    Dim tmp As MSForms.Label
    Set tmp = Me.Controls.Add("Forms.Label.1", TMP_NAME, False)
    tmp.Font.Name = fontName
    FontExists = (StrComp(tmp.Font.Name, fontName, vbTextCompare) = 0)
    Me.Controls.Remove TMP_NAME
End Function

Private Sub btnSaveChoices_Click()
    SaveUserChoices _
        Me.txtDefinitionsCol.Text, _
        Me.txtClauseCols.Text, _
        Me.txtRowRanges.Text

    MsgBox "Choices saved. They will auto-load next time.", vbInformation
End Sub

Private Sub btnRunRobertsMacro_Click()
    RunDefinitionsTool _
        Me.txtDefinitionsCol.Text, _
        Me.txtClauseCols.Text, _
        Me.txtRowRanges.Text
End Sub

Private Sub btnReportBug_Click()
    Dim mailTo As String
    Dim subj As String
    Dim body As String

    mailTo = "robert.stevens@freshfields.com"
    subj = "Robert's Definitions Tool - Bug Report"
    body = "Describe the issue here..."

    Application.FollowHyperlink _
        "mailto:" & mailTo & _
        "?subject=" & URLEncode(subj) & _
        "&body=" & URLEncode(body)
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
