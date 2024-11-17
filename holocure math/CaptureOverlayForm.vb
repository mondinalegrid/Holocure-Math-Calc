Imports System.Runtime.InteropServices

Public Class CaptureOverlayForm
    Inherits Form

    ' Import necessary function to change window style
    <DllImport("user32.dll")>
    Private Shared Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As Integer
    End Function

    Private captureRect As Rectangle
    Private dragStart As Point
    Private isDragging As Boolean = False
    Private moveStep As Integer = 10
    Friend WithEvents TextBox1 As TextBox
    Private resizeStep As Integer = 10

    ' Constants for window styles
    Private Const GWL_EXSTYLE As Integer = -20
    Private Const WS_EX_LAYERED As Integer = &H80000
    Private Const WS_EX_TRANSPARENT As Integer = &H20

    Public Sub New()
        ' Set up the form as a transparent overlay
        Me.FormBorderStyle = FormBorderStyle.None
        Me.BackColor = Color.DarkViolet
        Me.TransparencyKey = Color.DarkViolet ' Make this color transparent
        Me.Opacity = 0.5 ' Make the form semi-transparent
        Me.TopMost = True ' Ensure the form is always on top
        Me.ShowInTaskbar = False
        Me.Bounds = Screen.PrimaryScreen.Bounds ' Start as full screen

        ' Initialize the rectangle (set an initial size and location)
        captureRect = New Rectangle(800, 280, 700, 40) ' Initial area to capture

        ' Hook up the mouse events
        AddHandler Me.MouseDown, AddressOf OverlayForm_MouseDown
        AddHandler Me.MouseMove, AddressOf OverlayForm_MouseMove
        AddHandler Me.MouseUp, AddressOf OverlayForm_MouseUp

        ' Handle KeyDown for arrow key movement and resizing
        Me.KeyPreview = True
        AddHandler Me.KeyDown, AddressOf OverlayForm_KeyDown
    End Sub

    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)

        ' Set the form to be TopMost (always on top)
        Me.TopMost = True

        ' Set the form to be click-through
        Dim exStyle As Integer = GetWindowLong(Me.Handle, GWL_EXSTYLE)
        SetWindowLong(Me.Handle, GWL_EXSTYLE, exStyle Or WS_EX_LAYERED Or WS_EX_TRANSPARENT)
    End Sub

    ' Mouse events (same as before)
    Private Sub OverlayForm_MouseDown(sender As Object, e As MouseEventArgs)
        If captureRect.Contains(e.Location) Then
            dragStart = e.Location
            isDragging = True
        End If
    End Sub

    Private Sub OverlayForm_MouseMove(sender As Object, e As MouseEventArgs)
        If isDragging Then
            Dim width As Integer = e.X - dragStart.X
            Dim height As Integer = e.Y - dragStart.Y
            captureRect = New Rectangle(dragStart.X, dragStart.Y, width, height)
            Me.Invalidate()
        End If
    End Sub

    Private Sub OverlayForm_MouseUp(sender As Object, e As MouseEventArgs)
        isDragging = False
    End Sub

    Private Sub OverlayForm_KeyDown(sender As Object, e As KeyEventArgs)
        ' Handle arrow keys for movement
        Select Case e.KeyCode
            Case Keys.Up
                captureRect.Y -= moveStep
            Case Keys.Down
                captureRect.Y += moveStep
            Case Keys.Left
                captureRect.X -= moveStep
            Case Keys.Right
                captureRect.X += moveStep
            Case Keys.Oemplus ' '+' key
                captureRect.Height += resizeStep
            Case Keys.OemMinus ' '-' key
                captureRect.Height -= resizeStep
            Case Keys.OemOpenBrackets ' '[' key
                captureRect.Width -= resizeStep
            Case Keys.OemCloseBrackets ' ']' key
                captureRect.Width += resizeStep
        End Select

        ' Redraw the overlay form
        Me.Invalidate()
    End Sub

    ' Paint the capture area (same as before)
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Using brush As New SolidBrush(Color.FromArgb(100, 255, 0, 0))
            e.Graphics.FillRectangle(brush, captureRect)
        End Using
        MyBase.OnPaint(e)
    End Sub

    Public Function CaptureImage() As Bitmap
        ' Capture the selected area into an in-memory Bitmap
        Dim screenCapture As New Bitmap(captureRect.Width, captureRect.Height)
        Using g As Graphics = Graphics.FromImage(screenCapture)
            g.CopyFromScreen(captureRect.Left, captureRect.Top, 0, 0, captureRect.Size)
        End Using

        Return screenCapture
    End Function
End Class