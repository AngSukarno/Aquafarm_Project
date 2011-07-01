
Public Class WindowWrapper

    Implements System.Windows.Forms.IWin32Window
    Private _hwnd As IntPtr

    Public Sub New(ByVal handle As IntPtr)
        _hwnd = handle
    End Sub

    Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
        Get
            Return _hwnd
        End Get
    End Property

End Class

