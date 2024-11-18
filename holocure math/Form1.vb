Imports Emgu.CV
Imports Emgu.CV.CvEnum
Imports Emgu.CV.BitmapExtension
Imports Emgu.CV.Structure
Imports System.Runtime.InteropServices
Imports System.Drawing.Imaging

Public Class Form1

#Region "Variables"
    Private image As Mat
    Private templates As New Dictionary(Of String, Mat)
    Private windowRegions As RECT
    Private windowHandle As IntPtr
    Private isRunning As Boolean = False
    Private loopThread As System.Threading.Thread
#End Region

#Region "Form Events"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.ScrollBars = ScrollBars.Vertical
        TextBox1.ReadOnly = True
        MaximizeBox = False
        FormBorderStyle = FormBorderStyle.FixedDialog
        Button2.Text = "Capture"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If isRunning Then
            isRunning = False
            Button2.Text = "Capture"
        Else
            isRunning = True
            Button2.Text = "Stop"
            loopThread = New System.Threading.Thread(AddressOf RunLoop)
            loopThread.Start()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim customText = TextBox2.Text
        Dim evaluatedResults As String = ProcessMathExpressions(customText)

        TextBox1.Text = TextBox1.Text & Environment.NewLine &
                        "Extracted Expression: " & customText & Environment.NewLine &
                        "Evaluated Results: " & evaluatedResults
        TextBox1.SelectionStart = TextBox1.Text.Length
        TextBox1.ScrollToCaret()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

#End Region

#Region "Sub/Functions"
    Private Function TemplateMatching(bitmapImage As Bitmap) As String
        ' Load the main image
        image = ToMat(bitmapImage)

        ' Convert image to grayscale
        Dim grayImage As Mat = New Mat()
        CvInvoke.CvtColor(image, grayImage, ColorConversion.Bgr2Gray)
        CvInvoke.Threshold(image, image, 127, 255, ThresholdType.Binary)

        Dim path_to_file = Application.StartupPath & "templates\"

        templates.Clear()
        ' Load the template images
        templates.Add("0", CvInvoke.Imread($"{path_to_file}0.png", ImreadModes.Grayscale))
        templates.Add("1", CvInvoke.Imread($"{path_to_file}1.png", ImreadModes.Grayscale))
        templates.Add("2", CvInvoke.Imread($"{path_to_file}2.png", ImreadModes.Grayscale))
        templates.Add("3", CvInvoke.Imread($"{path_to_file}3.png", ImreadModes.Grayscale))
        templates.Add("4", CvInvoke.Imread($"{path_to_file}4.png", ImreadModes.Grayscale))
        templates.Add("5", CvInvoke.Imread($"{path_to_file}5.png", ImreadModes.Grayscale))
        templates.Add("6", CvInvoke.Imread($"{path_to_file}6.png", ImreadModes.Grayscale))
        templates.Add("7", CvInvoke.Imread($"{path_to_file}7.png", ImreadModes.Grayscale))
        templates.Add("8", CvInvoke.Imread($"{path_to_file}8.png", ImreadModes.Grayscale))
        templates.Add("9", CvInvoke.Imread($"{path_to_file}9.png", ImreadModes.Grayscale))
        templates.Add(")", CvInvoke.Imread($"{path_to_file}closedp.png", ImreadModes.Grayscale))
        templates.Add("/", CvInvoke.Imread($"{path_to_file}divide.png", ImreadModes.Grayscale))
        templates.Add("*", CvInvoke.Imread($"{path_to_file}multiply.png", ImreadModes.Grayscale))
        templates.Add("(", CvInvoke.Imread($"{path_to_file}openp.png", ImreadModes.Grayscale))
        templates.Add("+", CvInvoke.Imread($"{path_to_file}plus.png", ImreadModes.Grayscale))
        templates.Add("?", CvInvoke.Imread($"{path_to_file}question.png", ImreadModes.Grayscale))
        templates.Add("=", CvInvoke.Imread($"{path_to_file}equals.png", ImreadModes.Grayscale))
        templates.Add("-", CvInvoke.Imread($"{path_to_file}minus.png", ImreadModes.Grayscale))

        ' Perform template matching
        Return ExtractExpressionInOrder(grayImage)
    End Function

    ' Extract expression in order based on the x-position of the match
    Private Function ExtractExpressionInOrder(mainImage As Mat) As String
        Dim extractedText As New List(Of String)()
        Dim matchResults As New List(Of (String, Point))()

        ' Perform template matching for each template
        For Each templatePair In templates
            Dim template = templatePair.Value
            Dim result As New Mat()

            ' Match the template in the main image
            CvInvoke.MatchTemplate(mainImage, template, result, TemplateMatchingType.CcoeffNormed)

            ' Iterate through all matches above a confidence threshold
            Dim threshold As Double = 0.8 ' Adjust as needed
            Dim minDistance As Integer = 10 ' Minimum distance to consider as non-duplicate
            While True
                Dim minVal As Double, maxVal As Double
                Dim minLoc As Point, maxLoc As Point
                CvInvoke.MinMaxLoc(result, minVal, maxVal, minLoc, maxLoc)

                ' Check if the match is above the threshold
                If maxVal >= threshold Then
                    ' Check for duplicates
                    Dim isDuplicate As Boolean = matchResults.Any(Function(c) Math.Abs(c.Item2.X - maxLoc.X) < minDistance AndAlso Math.Abs(c.Item2.Y - maxLoc.Y) < minDistance)

                    If Not isDuplicate Then
                        If templatePair.Key = "-" Then
                            Dim matchedWidth As Integer = template.Width
                            Dim matchedHeight As Integer = template.Height
                            Dim aspectRatio As Double = matchedWidth / matchedHeight
                            Dim toleranceWidth As Integer = template.Width * 0.2
                            Dim toleranceHeight As Integer = template.Height * 0.2
                            Dim minAspectRatio As Double = aspectRatio * 0.8
                            Dim maxAspectRatio As Double = aspectRatio * 1.2

                            If Math.Abs(template.Width - matchedWidth) < toleranceWidth AndAlso
                               Math.Abs(template.Height - matchedHeight) < toleranceHeight AndAlso
                               aspectRatio > minAspectRatio AndAlso aspectRatio < maxAspectRatio Then
                                ' Ensure match is not near the bottom of a number
                                If Not IsNearNumberBottom(maxLoc, mainImage) Then
                                    matchResults.Add((templatePair.Key, maxLoc))
                                End If
                            End If
                        Else
                            matchResults.Add((templatePair.Key, maxLoc))
                        End If
                    End If

                    ' Suppress the current match to find the next one
                    CvInvoke.Rectangle(result, New Rectangle(maxLoc, template.Size), New MCvScalar(0), -1)
                Else
                    Exit While ' No more matches above the threshold
                End If
            End While
        Next

        ' Sort the match results by the x-coordinate (left to right)
        matchResults.Sort(Function(a, b) a.Item2.X.CompareTo(b.Item2.X))

        ' Combine adjacent matches based on proximity (group digits together)
        Dim currentGroup As String = ""
        Dim lastX As Integer = -1

        For Each match In matchResults
            Dim matchValue As String = match.Item1
            Dim matchX As Integer = match.Item2.X

            ' If the current match is far enough from the last match, treat it as a new group
            If lastX = -1 OrElse matchX > lastX + 20 Then ' Adjust threshold (e.g., 20 pixels) as needed
                ' Add the previous group to the result
                If currentGroup <> "" Then
                    extractedText.Add(currentGroup)
                End If
                currentGroup = matchValue ' Start a new group
            Else
                ' Otherwise, append the current match to the existing group
                currentGroup &= matchValue
            End If

            ' Update lastX to be the current match's x-coordinate
            lastX = matchX
        Next

        ' Add the last group to the result if necessary
        If currentGroup <> "" Then
            extractedText.Add(currentGroup)
        End If

        ' Join the extracted text to form the expression
        Return String.Join(" ", extractedText)
    End Function

    Private Function IsNearNumberBottom(location As Point, mainImage As Mat) As Boolean
        ' Set a tolerance for proximity to the bottom of detected numbers
        Dim proximityTolerance As Integer = 5 ' Adjust based on the size of characters in the image

        ' Iterate through possible number templates to find bounding boxes
        Dim path_to_file = Application.StartupPath & "templates\"
        Dim numberTemplates As List(Of String) = New List(Of String) From {$"{path_to_file}0.png", $"{path_to_file}1.png", $"{path_to_file}2.png", $"{path_to_file}3.png", $"{path_to_file}4.png", $"{path_to_file}5.png", $"{path_to_file}6.png", $"{path_to_file}7.png", $"{path_to_file}8.png", $"{path_to_file}9.png"}
        For Each templatePath In numberTemplates
            ' Load the number template
            Dim template As Mat = CvInvoke.Imread(templatePath, ImreadModes.Grayscale)

            ' Perform template matching
            Dim result As Mat = New Mat()
            CvInvoke.MatchTemplate(mainImage, template, result, TemplateMatchingType.CcorrNormed)

            Dim threshold As Double = 0.8 ' Confidence threshold for matching numbers
            While True
                Dim minVal As Double, maxVal As Double
                Dim minLoc As Point, maxLoc As Point
                CvInvoke.MinMaxLoc(result, minVal, maxVal, minLoc, maxLoc)

                ' Check if the match is above the threshold
                If maxVal >= threshold Then
                    ' Get the bounding box of the detected number
                    Dim detectedRect As Rectangle = New Rectangle(maxLoc, template.Size)

                    ' Check if the detected location is near the bottom of this number
                    Dim bottomY As Integer = detectedRect.Bottom
                    If Math.Abs(location.Y - bottomY) <= proximityTolerance AndAlso
                     location.X >= detectedRect.Left AndAlso location.X <= detectedRect.Right Then
                        Return True ' The location is near the bottom of a number
                    End If

                    If detectedRect.Contains(location) Then
                        Return True ' The location is inside or on the number
                    End If

                    ' Suppress the current match to find the next one
                    CvInvoke.Rectangle(result, detectedRect, New MCvScalar(0), -1)
                Else
                    Exit While ' No more matches above the threshold
                End If
            End While
        Next

        ' If no number regions are near the location, return false
        Return False
    End Function

    ' Process and evaluate math expressions
    Private Function ProcessMathExpressions(extractedText As String) As String
        ' Split the extracted text into individual expressions (assuming each expression is on a new line)
        Dim mathExpressions As String() = extractedText.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

        ' Process each expression and evaluate
        Dim results As New List(Of String)()

        For Each expression In mathExpressions
            ' Clean up the expression (remove extra spaces, replace 'x' with '*')
            expression = expression.Trim().Replace("x", "*")

            ' Evaluate the math expression
            Try
                Dim result As String = EvaluateMathExpression(expression)
                results.Add($"{result}")
            Catch ex As Exception
                results.Add($"Error evaluating expression '{expression}': {ex.Message}")
            End Try
        Next

        ' Return the evaluated results as a concatenated string
        Return String.Join(Environment.NewLine, results)
    End Function

    ' Helper function to evaluate a single math expression
    Private Function EvaluateMathExpression(expression As String) As String
        Try
            ' Use DataTable.Compute to evaluate the expression
            Dim result As Object = New DataTable().Compute(expression, Nothing)

            ' Return the result as a string
            Return result.ToString()
        Catch ex As Exception
            ' If there's an error in the expression, return the error message
            Return $"Error: {ex.Message}"
        End Try
    End Function

    Public Function GetHolocureDimensions() As Boolean
        windowHandle = FindWindow("YYGameMakerYY", "HoloCure")

        If windowHandle <> IntPtr.Zero Then
            ' Get the window's position and size
            If GetWindowRect(windowHandle, windowRegions) Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Public Function CaptureWindow() As Bitmap
        Dim screenCapture As New Bitmap(windowRegions.Width, 40)
        Using g As Graphics = Graphics.FromImage(screenCapture)
            g.CopyFromScreen(windowRegions.Left + 296, windowRegions.Top + 117, 0, 0, New Size(windowRegions.Width - 670, 40))
        End Using
        Return screenCapture
    End Function

    Private Sub RunLoop()
        Dim prevExtract As String = ""
        While isRunning
            Me.Invoke(Sub()
                          If GetHolocureDimensions() Then
                              Dim captureImage = CaptureWindow()
                              Dim extractedText = TemplateMatching(captureImage).Replace("-2", "2").Replace("=", "").Replace("?", "").Trim
                              Dim evaluatedResults As String = ProcessMathExpressions(extractedText)

                              TextBox1.Text = "Extracted Expression: " & extractedText & Environment.NewLine &
                                              "Evaluated Results: " & evaluatedResults
                              TextBox1.SelectionStart = TextBox1.Text.Length
                              TextBox1.ScrollToCaret()

                              If prevExtract <> extractedText Then
                                  TextBox2.Text = extractedText
                              End If
                          Else
                              isRunning = False
                              MessageBox.Show("Holocure is minimized or closed")
                          End If
                      End Sub)

            System.Threading.Thread.Sleep(800)
        End While

        Me.Invoke(Sub() TextBox1.Text = "Stopped")
    End Sub
#End Region

#Region "WinAPI"
    ' Struct to store window's position and size
    Public Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer

        ' Calculate Width and Height
        Public ReadOnly Property Width As Integer
            Get
                Return Right - Left
            End Get
        End Property

        Public ReadOnly Property Height As Integer
            Get
                Return Bottom - Top
            End Get
        End Property
    End Structure

    ' API declaration for GetWindowRect
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetWindowRect(ByVal hwnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function

    ' API declaration for FindWindow
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    End Function
#End Region


End Class
