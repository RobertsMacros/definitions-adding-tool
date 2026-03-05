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
    '=== Load saved preferences ===
    Dim defsCol As String, clauseCols As String, rowRanges As String
    LoadUserChoices defsCol, clauseCols, rowRanges
    '=== Set labels and populate textboxes ===
    Me.lblDefsCol.Caption = "Definitions column (e.g. AL):"
    Me.lblClauseCols.Caption = "Clause columns (e.g. H:AK):"
    Me.lblRowRanges.Caption = "Rows to process (e.g. 2:500):"
    If defsCol <> "" Then Me.txtDefinitionsCol.Text = defsCol
    If clauseCols <> "" Then Me.txtClauseCols.Text = clauseCols
    If rowRanges <> "" Then Me.txtRowRanges.Text = rowRanges
    Me.btnRunRobertsMacro.Caption = "Run Macro"
    Me.btnSaveChoices.Caption = "Save Choices"
    Me.btnClose.Caption = "Close"

End Sub
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
Private Sub btnClose_Click()
    Unload Me
End Sub
