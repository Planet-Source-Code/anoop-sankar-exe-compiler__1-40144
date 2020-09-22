VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picInfo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   960
      ScaleHeight     =   2715
      ScaleWidth      =   5475
      TabIndex        =   6
      Top             =   1088
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "$$"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   4
         Left            =   2160
         TabIndex        =   18
         Top             =   2280
         Width           =   210
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "$$"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2160
         TabIndex        =   17
         Top             =   1920
         Width           =   210
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "$$"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2160
         TabIndex        =   16
         Top             =   1560
         Width           =   210
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "$$"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   15
         Top             =   1200
         Width           =   210
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "$$"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   14
         Top             =   840
         Width           =   210
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Overhead % :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   13
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Size :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1020
         TabIndex        =   12
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Exe Overhead  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   510
         TabIndex        =   11
         Top             =   1560
         Width           =   1545
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "File Size  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1050
         TabIndex        =   10
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "File Name :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Exe Information"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5280
         TabIndex        =   7
         Top             =   0
         Width           =   150
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   30
         X2              =   5417
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         Index           =   2
         X1              =   30
         X2              =   5417
         Y1              =   2655
         Y2              =   2655
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   30
         X2              =   5200
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   30
         X2              =   5200
         Y1              =   135
         Y2              =   135
      End
   End
   Begin VB.CommandButton cmdExeInfo 
      Caption         =   "&EXE Info"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmMain.frx":1042
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "www.smilehouse.cjb.net"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   5
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Coded by Anoop Sankar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   4
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Created using prjExeCompiler"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   3
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Compiled Exe File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   4200
      X2              =   7320
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   4200
      X2              =   7320
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Height          =   2340
      Left            =   4200
      Top             =   120
      Width           =   3150
   End
   Begin VB.Image imgPic 
      Height          =   2340
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3150
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------
'       Copyright 2002, Anoop Sankar
'You may freely use, modify and distribute this source
'code, provided that you do not remove this message.
'But, you are NOT allowed to distribute the compiled
'version (.EXE,.DLL,.OCX etc etc.) of this program
'or any program which uses the below code without my
'consent.
'
'If you modified something, put your name below..
'
'Orginal Code : Anoop Sankar (anoops@gmx.net)
'Modified by  : No one so far
'
'Last Update : Oct 25,2002
'Visit www.smilehouse.cjb.net for more source code
'------------------------------------------------------
'
'This is the source for the rdfce.ext file,
'which is used as the template for creating the exe.
'
'This project will not run properly in the IDE.
'Check readme.html for more information
'
'------------------------------------------------------

Dim PropBag As New PropertyBag      'the property bag

Private Sub Form_Load()
    'On Local Error Resume Next
    
    Dim BeginPos As Long
    Dim varTemp As Variant
    
    Dim byteArr() As Byte
    
    Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1
        Get #1, LOF(1) - 3, BeginPos    'get the start position of data

        Seek #1, BeginPos               'seek to data start
        Get #1, , varTemp               'get property bag contents
        
        byteArr = varTemp
        PropBag.Contents = byteArr      'load property bag
    
        PropBag.WriteProperty "LOF", LOF(1) 'a few extra props
        PropBag.WriteProperty "BeginPos", BeginPos
    Close #1
        
    
    'password protection
    'I know that this is not tight, but just for a demo
    If Val(PropBag.ReadProperty("Protected", "0")) > 0 Then
        Dim PassInp As String
        
        PassInp = InputBox("Enter password:", "Password Required")
        
        If PassInp <> PropBag.ReadProperty("Password") Then
            MsgBox "Password not valid", vbCritical, "Nice Try!"
            End
        End If
    End If
    
    With PropBag
        txtText.Text = .ReadProperty("Text")
        Set imgPic.Picture = .ReadProperty("Picture")
        Me.Caption = .ReadProperty("Caption")
    End With

End Sub

Private Sub cmdExeInfo_Click()
    'display exe stats
    
    lblInfo(0).Caption = App.EXEName & ".exe"
    lblInfo(1).Caption = PropBag.ReadProperty("LOF") & " bytes"
    lblInfo(2).Caption = PropBag.ReadProperty("BeginPos") & " bytes"
    lblInfo(3).Caption = (PropBag.ReadProperty("LOF") - PropBag.ReadProperty("BeginPos")) & " bytes"
    lblInfo(4).Caption = Format((PropBag.ReadProperty("BeginPos") / PropBag.ReadProperty("LOF")) * 100, "0.00") & " %"

    picInfo.Visible = True
End Sub

Private Sub Label3_Click()
    picInfo.Visible = False
End Sub
