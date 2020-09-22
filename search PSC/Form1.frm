VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10635
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":030A
      Left            =   8760
      List            =   "Form1.frx":032F
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   255
      Left            =   8760
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "get html"
      Height          =   255
      Left            =   8760
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8760
      TabIndex        =   3
      Text            =   "¾ ßúzzèÐ"
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   8760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4815
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8655
      ExtentX         =   15266
      ExtentY         =   8493
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hw As Long
Dim sHtml As String
Dim iCase As Integer
Dim URL As String

Private Sub Command1_Click()
Text1.Text = ""
    iCase = 2

If Combo1.Text = ".NET" Then URL = "http://www.planet-source-code.com/vb/default.asp?lngWId=10"
If Combo1.Text = "ASP/VBScript" Then URL = "http://www.planet-source-code.com/vb/default.asp?lngWId=4"
If Combo1.Text = "C/C++" Then URL = "http://www.planet-source-code.com/vb/default.asp?lngWId=3"
If Combo1.Text = "Cold Fusion" Then URL = "http://www.planet-source-code.com/vb/default.asp?lngWId=9"
If Combo1.Text = "Delphi" Then URL = "http://www.planet-source-code.com/vb/default.asp?lngWId=7"
If Combo1.Text = "Java/Javascript" Then URL = "http://www.planet-source-code.com/vb/default.asp?lngWId=2"
If Combo1.Text = "LISP" Then URL = "http://www.planet-source-code.com/vb/default.asp?lngWId=13"
If Combo1.Text = "Perl" Then URL = "http://www.planet-source-code.com/vb/default.asp?lngWId=6"
If Combo1.Text = "PHP" Then URL = "http://www.planet-source-code.com/vb/default.asp?lngWId=8"
If Combo1.Text = "SQL" Then URL = "http://www.planet-source-code.com/vb/default.asp?lngWId=5"
If Combo1.Text = "Visual Basic" Then URL = "http://www.planet-source-code.com/vb/default.asp?lngWId=1"

WebBrowser1.Navigate URL
End Sub

Private Sub Command2_Click()
Text1.Text = ""
    iCase = 1
    WebBrowser1.Navigate "http://www.planet-source-code.com/vb/default.asp?lngWId=1"
End Sub


Private Sub Form_Load()
Combo1.Text = "Visual Basic"
    WebBrowser1.Navigate2 "about:<html><body bgcolor=""000000""><center><font size=8><b><font color=""red"">Search - Planet Source Code<BR><font size=3>By: <font color=""#FFCC66""><B><strong>Underground Technologies Inc.</strong></b><p align=""center""><font size=""1"">Copyright 2000-2008 Underground Technologies, Inc.&nbsp; All Rights Reserved.<br></font></body></html>"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'SetWindowLong hw, GWL_WNDPROC, origWndProc
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Form1 = Nothing
End
End Sub


Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Select Case iCase
  Case 1
        iCase = 0
        sHtml = WebBrowser1.Document.All.Item(0).innerHTML
        Text1.Text = Text1.Text & sHtml & vbCrLf
        Text1.Text = Text1.Text & "===============" & vbCrLf
        iCase = 0
  Case 2
        With WebBrowser1.Document
            .All("txtCriteria").Value = Text2.Text
                Text1.Text = Text1.Text & "Filled Form" & vbCrLf
            .All("B1").Click
                Text1.Text = Text1.Text & "Button Clicked" & vbCrLf
                Text1.Text = Text1.Text & "===============" & vbCrLf
        End With
        iCase = 0
        URL = ""
End Select
End Sub
