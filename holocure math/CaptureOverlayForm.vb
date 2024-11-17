Public Class CaptureOverlayForm
    Inherits Form

    Private captureRect As Rectangle
    Private dragStart As Point
    Private isDragging As Boolean = False
    Private moveStep As Integer = 10
    Friend WithEvents TextBox1 As TextBox
    Private resizeStep As Integer = 10

    Public Sub New()
        ' Set up the form as a transparent overlay
        Me.FormBorderStyle = FormBorderStyle.None
        Me.BackColor = Color.LimeGreen ' Pick a color that won't conflict with the background
        Me.TransparencyKey = Color.LimeGreen ' Make this color transparent
        Me.Opacity = 0.5 ' Make the form semi-transparent
        Me.TopMost = True ' Ensure the form is always on top
        Me.ShowInTaskbar = False
        Me.Bounds = Screen.PrimaryScreen.Bounds ' Start as full screen

        ' Initialize the rectangle (set an initial size and location)
        captureRect = New Rectangle(800, 280, 200, 60) ' Initial area to capture

        ' Hook up the mouse events
        AddHandler Me.MouseDown, AddressOf OverlayForm_MouseDown
        AddHandler Me.MouseMove, AddressOf OverlayForm_MouseMove
        AddHandler Me.MouseUp, AddressOf OverlayForm_MouseUp

        ' Handle KeyDown for arrow key movement and resizing
        Me.KeyPreview = True
        AddHandler Me.KeyDown, AddressOf OverlayForm_KeyDown
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

    Private Sub InitializeComponent()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(139, 26)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 23)
        Me.TextBox1.TabIndex = 0
        '
        'CaptureOverlayForm
        '
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "CaptureOverlayForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
End Class