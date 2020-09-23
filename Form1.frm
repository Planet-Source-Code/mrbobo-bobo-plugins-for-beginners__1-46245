VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2250
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2190
      ScaleWidth      =   2760
      TabIndex        =   4
      Top             =   1620
      Width           =   2820
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   4500
      Pattern         =   "*.dll"
      TabIndex        =   3
      Top             =   3060
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRoot3 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Text            =   "Root3"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtRoot2 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "Root2"
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox txtRoot1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Root"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Menu mnuPlugInsBase 
      Caption         =   "PlugIns"
      Begin VB.Menu mnuPlugIns 
         Caption         =   "Empty"
         Enabled         =   0   'False
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2003********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive


'PlugIns for beginners
'What is a PlugIn ?
'A PlugIn is a dll/exe/script that is installed separately from an application
'that can perform tasks within that application. It should be able to pass objects freely
'between the application and the plugin.

'In this example we have an executible containing a class module - this is the object we
'will throw about between the executible and the plugin. The plugins are ActiveX dlls.

'There are many ways of achieving this functionality. This example is a very simple
'method - that's why I used it here as a demonstration.



'**********IMPORTANT _ READ FIRST ***********************************************
'To run this demo you need to compile Project2 and Project3 - these are your Plugins
'You should compile them to the same directory as Project1 - that's where this
'app will look for any Plugins.
'********************************************************************************
Option Explicit
Dim mR As ClRoot

Private Sub Form_Load()
    'Load our class and fill it with some data
    Set mR = New ClRoot
    mR.Root1 = txtRoot1.Text
    mR.Root2 = txtRoot2.Text
    mR.Root3 = txtRoot3.Text
    Set mR.BoboPic = Picture1
    'Use the App.path as our Plugin directory
    'Use a filebox with it's pattern set to "dll" to get any Plugins
    File1.Path = App.Path
    'Load up the PlugIn menu
    BuildPlugInMenu
End Sub
Private Sub mnuPlugIns_Click(Index As Integer)
    Dim mPlug As Object
    'Load the Plugin
    Set mPlug = CreateObject(mnuPlugIns(Index).Caption & ".ClPlugIn")
    'Pass our class to the Plugin and get it to make some changes
    mPlug.LoadClass mR
    'Show the changes
    txtRoot1.Text = mR.Root1
    txtRoot2.Text = mR.Root2
    txtRoot3.Text = mR.Root3
    'Clean up - we're finished with the Plugin
    Set mPlug = Nothing
End Sub

Private Sub txtRoot1_Change()
    mR.Root1 = txtRoot1.Text
End Sub

Private Sub txtRoot2_Change()
    mR.Root2 = txtRoot2.Text
End Sub

Private Sub txtRoot3_Change()
    mR.Root3 = txtRoot3.Text
End Sub

Public Sub BuildPlugInMenu()
    'Plugins must be in a format we can recognize
    'Photoshop for example has a set format for Plugins
    'as do we for this application. Our "format" says that
    'Plugins should have a class called "ClPlugin"
    'These classes should be aware of the make-up of
    'objects passed to them
    Dim z As Long
    Dim ValidObjects As Collection
    Dim ValidName As String
    If File1.ListCount > 0 Then
        Set ValidObjects = New Collection
        For z = 0 To File1.ListCount - 1
            'Make sure we can load the dll
            ValidName = Left(File1.List(z), Len(File1.List(z)) - 4)
            If CanLoadObject(ValidName & ".ClPlugIn") Then ValidObjects.Add ValidName
        Next
        If ValidObjects.Count > 0 Then
            'For all the loadable dll's create a menu
            For z = 1 To ValidObjects.Count
                If z > 1 Then Load mnuPlugIns(z - 1)
                mnuPlugIns(z - 1).Caption = ValidObjects(z)
                mnuPlugIns(z - 1).Enabled = True
                mnuPlugIns(z - 1).Visible = True
            Next
        End If
    Else
        'just in case you forgot to compile the dll's !!
        MsgBox "To run this demo you need to compile Project2 and Project3 - these are your Plugins." & vbCrLf & _
        "You should compile them to the same directory as Project1 - that's where this" & vbCrLf & _
        "app will look for any Plugins."
        Unload Me
    End If

End Sub

Public Function CanLoadObject(mObject As String) As Boolean
    Dim mPlug As Object
    'If we can't load it we'll end up at Woops
    On Error GoTo Woops
    Set mPlug = CreateObject(mObject)
    Set mPlug = Nothing
    CanLoadObject = True
    Exit Function
Woops:
    Set mPlug = Nothing
End Function
