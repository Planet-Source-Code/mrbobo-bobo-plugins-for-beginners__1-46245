VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClPlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
'***************Copyright PSST 2003********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive
Option Explicit

Public Function LoadClass(mClass As Object)
    'mClass is ClRoot in Project1
    'We're just passing strings nd a Picturebox but they could be
    'pictures that we're performing a filter function
    'or anything that a Class module may contain
    mClass.Root1 = UCase(mClass.Root1)
    mClass.Root2 = UCase(mClass.Root2)
    mClass.Root3 = UCase(mClass.Root3)
    mClass.BoboPic.BackColor = vbRed
End Function
 
