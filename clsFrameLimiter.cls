VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFrameLimiter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Thanks To Geoffrey Hazen For This Class

Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private m_CurFrequency As Currency
Private m_HasCounter As Boolean
Private m_FrameStart As Currency
Private m_FrameEnd As Currency
Private m_CurTime As Currency
Private m_Delay As Currency

Private Sub Class_Initialize()
  'Get Performance Frequency when class is Initialized.
  m_HasCounter = QueryPerformanceFrequency(m_CurFrequency)
  m_CurFrequency = m_CurFrequency * 10000
End Sub

Public Sub LimitFrames(ByVal nFPS As Integer)
  'Get current counter and set as Frame End.
  QueryPerformanceCounter m_FrameEnd
  
  'Calculate delay by subtracting Frame Start from Frame End from Disired FPS.
  m_Delay = ((1000 / nFPS) * m_CurFrequency / 10000000) - (m_FrameEnd - m_FrameStart)
  
  'Loop while surrendering to other apps until desired delay time has been met.
  Do
    DoEvents
    QueryPerformanceCounter m_CurTime
  Loop Until (m_CurTime - m_FrameEnd) >= m_Delay
  
  'Get current counter and set as Frame Start.
  QueryPerformanceCounter m_FrameStart
End Sub
