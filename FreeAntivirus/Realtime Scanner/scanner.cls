VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scanner"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Implements IShellExecuteHookW
Private m_pOldIShellExecuteHookW As Long


Private Sub Class_Initialize()
    Dim pShellExecuteHookW As IShellExecuteHookW
    Set pShellExecuteHookW = Me
    
    m_pOldIShellExecuteHookW = SwapVtableEntry(ObjPtr(pShellExecuteHookW), 4, AddressOf Execute)
End Sub

Public Sub IShellExecuteHookW_Execute(pei As olelib.SHELLEXECUTEINFO)
    'Execute
End Sub

Private Sub Class_Terminate()
   Dim pShellExecuteHookW As IShellExecuteHookW
   Set pShellExecuteHookW = Me
   m_pOldIShellExecuteHookW = SwapVtableEntry(ObjPtr(pShellExecuteHookW), 4, m_pOldIShellExecuteHookW)
End Sub

