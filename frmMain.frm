VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compiler"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog ComDLG 
      Left            =   6120
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame 
      Caption         =   "Exe Details"
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtPass 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Text            =   "password"
         Top             =   3600
         Width           =   2295
      End
      Begin VB.CheckBox chkPass 
         Caption         =   "Password Protect with"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "Enter caption here..."
         Top             =   360
         Width           =   6375
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
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmMain.frx":0000
         Top             =   720
         Width           =   3135
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  'Dot
         Height          =   2340
         Left            =   3360
         Top             =   720
         Width           =   3150
      End
      Begin VB.Label lblAddPic 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click to change picture"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4065
         TabIndex        =   6
         Top             =   3120
         Width           =   1665
      End
      Begin VB.Image imgPic 
         Height          =   2340
         Left            =   3360
         Picture         =   "frmMain.frx":0015
         Stretch         =   -1  'True
         Top             =   720
         Width           =   3150
      End
   End
   Begin VB.TextBox txtExeFile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "test.exe"
      Top             =   4530
      Width           =   5295
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "&Compile"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   2100
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   2
      Left            =   4680
      TabIndex        =   9
      Top             =   4920
      Width           =   2070
   End
   Begin VB.Label Label 
      Caption         =   "Exe File Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
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
'-------------------------------------------------------
'
'Purpose of the project is to create a stand alone exe
'file from VB code. It doesn't do this directly, but uses
'a simple work around. I think this is quite a useful
'way to do this. If you think otherwise or if you have
'other methods, I would love to hear from you.
'
'The method is .. write to the end of a pre-created exe
'file, in binary mode.
'
'Check 'readme.html' for more details.
'
'-------------------------------------------------------


Private Sub cmdCompile_Click()

    'This is were all the action takes place

    On Local Error GoTo errTrap

    Dim BeginPos As Long            'variable to store the start of data
    Dim PropBag As New PropertyBag  'property bag to store the data
    Dim varTemp As Variant          'for file writing
    
    'Below section loads data into the property bag.
    With PropBag
        .WriteProperty "Caption", txtCaption.Text
        .WriteProperty "Text", txtText.Text
        .WriteProperty "Picture", imgPic.Picture
        .WriteProperty "Protected", chkPass.Value
        .WriteProperty "Password", txtPass.Text
        'You may add your own propery using the syntax
        'PropBag.WriteProperty "<property name>",<property value>
        'As you might have noticed, property value can be anything
        'string, picture, or numerical.
    End With
    
    'rdfce.ext is the template we use to create our exe.
    'RDFCE = Renamed Dummy For Creating Executable ;-)
    '(source of that file is included too)
    
    'first copy that file to the user provided file name.
    FileCopy App.Path & "\rdfce.ext", App.Path & "\" & txtExeFile.Text
    
    'now open the file in binary mode
    Open App.Path & "\" & txtExeFile.Text For Binary As #1
        BeginPos = LOF(1)   'the point were we add extra data
                
        varTemp = PropBag.Contents
                
        Seek #1, LOF(1)
        Put #1, , varTemp   'write data
        Put #1, , BeginPos  'write starting point of extra data
    
    Close #1

    MsgBox "Exe File created without a problem", vbInformation, "Compilation Done"
    Exit Sub

    'Thats it! The exe is compiled.
    'Read the prjExeDummy (prjRdfce.vbp) to find how the
    'compiled exe works.

errTrap:
    'to err is electronic
    Msg = "There was an error during compilation" & vbCrLf
    Msg = Msg & vbCrLf & Err.Description
    MsgBox Msg, vbCritical, "Error"
End Sub


Private Sub lblAddPic_Click()

    On Local Error GoTo errTrap
    
    ComDLG.CancelError = True
    ComDLG.ShowOpen

    imgPic.Picture = LoadPicture(ComDLG.FileName)

errTrap:

End Sub

Private Sub chkPass_Click()
    
    On Local Error Resume Next
    
    If chkPass.Value > 0 Then
        txtPass.Enabled = True
        txtPass.SetFocus
    Else
        txtPass.Enabled = False
    End If
    
End Sub
