VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "lCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private m_HasCounter As Boolean
Private m_Freq As Currency
Private m_Start As Currency
Private m_Stop As Currency
Private m_Elapsed As Currency

Private Sub Class_Initialize()
  m_HasCounter = QueryPerformanceFrequency(m_Freq)
End Sub

Public Sub StartTimer()
  QueryPerformanceCounter m_Start
End Sub

Public Sub StopTimer()
  QueryPerformanceCounter m_Stop
End Sub

Public Property Get Elasped() As String
  Elasped = (((m_Stop - m_Start) / m_Freq) * 1000)
End Property


