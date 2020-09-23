VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "String test"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkTime 
      Caption         =   "Using FastRepace: "
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   18
      Top             =   7560
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox chkTime 
      Caption         =   "Using InStr combined with MID:"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   17
      Top             =   7200
      Width           =   2535
   End
   Begin VB.Frame frameOptions 
      Caption         =   "Options:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   7455
      Begin VB.TextBox txtReplaceWith 
         Height          =   285
         Left            =   3120
         TabIndex        =   14
         Text            =   "_"
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   3120
         TabIndex        =   12
         Text            =   "A"
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label2 
         Caption         =   "Replace with:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblFind 
         Caption         =   "Find what:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame frameResult 
      Caption         =   "Result:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   7455
      Begin VB.CheckBox chkTime 
         Caption         =   "Using the replace-function:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkTime 
         Caption         =   "Using a for-loop:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblTime 
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   9
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label lblTime 
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   8
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label lblTime 
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   7
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label lblStringSize 
         Caption         =   "Size of string:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblSize 
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label lblTime 
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   4
         Top             =   1800
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   8280
      Width           =   3495
   End
   Begin VB.TextBox txtSample 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   7455
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   8280
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
   
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean

   On Local Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)

End Function

Public Property Get IsIDE() As Boolean

    On Error GoTo Handler

    ' This code is only executed in the IDE - it raises an error if so
    Debug.Print 1 / 0
    
    Exit Property
Handler:

    ' We are executed from IDE
    IsIDE = True

End Property

Private Sub cmdExit_Click()

    ' Terminate application
    Unload Me

End Sub

Public Function TestMethod(lngNumber As Long, sText As String, sFind As String, sReplaceWith As String) As Double

    Dim objCount As New lCount, Tell As Long, Pos As Long

    ' Start the conting
    objCount.StartTimer

    Select Case lngNumber ' The different methods
    Case 0 ' For-loop
    
        ' Then, use the regular method
        For Tell = 1 To Len(sText)
            If Mid(sText, Tell, Len(sFind)) = sFind Then
                ' Replace character
                sText = Left(sText, Tell - 1) & sReplaceWith & Right(sText, Len(sText) - Tell)
            End If
        Next
    
    Case 1 ' Replace-method
    
        ' Replace text with the native replace-function
        sText = Replace(sText, sFind, sReplaceWith)
            
    Case 2 ' InStr
    
        ' Get the first position
        Pos = InStr(1, sText, sFind)
        
        Do Until Pos <= 0
        
            ' Replace that position
            sText = Left(sText, Pos - 1) & sReplaceWith & Right(sText, Len(sText) - Pos)
            
            ' Go to next
            Pos = InStr(Pos + 1, sText, sFind)
        
        Loop
        
    Case 3 ' VarPtrArray
    
        ' Replace characters
        sText = FastReplace(sText, sFind, sReplaceWith)
    
    End Select
    
    ' Stop it
    objCount.StopTimer
    
    ' Return the time
    TestMethod = objCount.Elasped

End Function

Private Sub cmdStart_Click()
    
    Dim Tell As Long, sText As String

    ' Show the lenght
    lblSize.Caption = Len(txtSample.Text)
    
    ' Show the result
    For Tell = lblTime.LBound To lblTime.UBound
        
        If chkTime(Tell).Value = 1 Then
            
            ' Get the string
            sText = txtSample.Text
            
            ' Replace the string
            lblTime(Tell) = Round(TestMethod(Tell, sText, txtFind.Text, txtReplaceWith.Text), 3) & " ms"
            
            ' Give the events a chance to run
            DoEvents
        
        End If
        
    Next

    ' Set the converted string
    txtSample.Text = sText

End Sub

Private Sub Form_Initialize()

    ' Neccessary for XP-style
    InitCommonControlsVB

End Sub

Private Sub Form_Load()

    ' See if the program is executed from the IDE
    If IsIDE Then
        MsgBox "Complie for better results", vbInformation
    End If

    ' Initializing the function, fixing it to not run badly the at the first
    ' executing due to delayed compiling. In this sample, this is not really required.
    FastReplace Space(128), " ", "  "

    ' Creating sample (neat, huh?)
    txtSample = FastReplace(Space(512), " ", String(66, "A") & vbCrLf)

End Sub
