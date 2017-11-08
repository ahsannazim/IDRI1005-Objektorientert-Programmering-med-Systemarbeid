'Option Strict On
'Option Explicit On
'Option Infer Off

'Imports System.Runtime.InteropServices

'Public Class PaintMessageHelper
'    Private Const WM_SETREDRAW As Integer = 11

'    'Public Declare Function SendMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal wMsg As Int32, ByVal wParam As Boolean, ByVal lParam As Int32) As Integer
'    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
'    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As Boolean, ByVal lParam As IntPtr) As IntPtr
'    End Function


'    '<DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
'    'Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
'    'End Function
'    Public Shared Sub SuspendDrawing(Target As Control)
'        'SendMessage(Target.Handle, WM_SETREDRAW, New IntPtr(0), New IntPtr(0))
'        SendMessage(Target.Handle, WM_SETREDRAW, False, New IntPtr(0))
'    End Sub
'    Public Shared Sub ResumeDrawing(Target As Control)
'        SendMessage(Target.Handle, WM_SETREDRAW, True, New IntPtr(0))
'        Target.Refresh()
'    End Sub
'End Class
