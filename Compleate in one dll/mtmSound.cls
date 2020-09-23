VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mtmSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public Function MakeSound(ByVal Freequancy As Long, ByVal Duration As Long) As Long
       MakeSound = Beep(Freequancy, Duration)
End Function


Public Function SoundForDetect()
    Beep 2500, 100
    Beep 2000, 50
    Beep 1000, 20
End Function

Public Function SoundForRemove()
 Beep 2400, 150
Beep 2050, 50
Beep 2230, 20
Beep 2800, 100
End Function
Public Function LongBeep()
 Beep 2400, 1000
End Function
Public Function ShortBeep()
 Beep 2400, 100
End Function
