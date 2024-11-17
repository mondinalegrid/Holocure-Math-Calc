Imports Emgu.CV
Imports Emgu.CV.CvEnum
Imports Emgu.CV.BitmapExtension
Imports Emgu.CV.Structure

Public Class Form1

    Private overlayForm As CaptureOverlayForm
    Private image As Mat
    Private templates As New Dictionary(Of String, Mat)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.ScrollBars = ScrollBars.Vertical
        TextBox1.ReadOnly = True
        MaximizeBox = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If overlayForm Is Nothing OrElse overlayForm.IsDisposed Then
            overlayForm = New CaptureOverlayForm()
            overlayForm.Show()
        Else
            overlayForm.BringToFront()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Capture the selected area
        If overlayForm IsNot Nothing Then
            Dim test = overlayForm.CaptureImage
            Dim extractedText = TemplateMatching(test).Replace("-2", "2").Replace("=", "").Replace("?", "").Trim
            Dim evaluatedResults As String = ProcessMathExpressions(extractedText)

            TextBox1.Text = TextBox1.Text & Environment.NewLine &
                            "Extracted Expression: " & extractedText & Environment.NewLine &
                            "Evaluated Results: " & evaluatedResults
            TextBox1.SelectionStart = TextBox1.Text.Length
            TextBox1.ScrollToCaret()

            TextBox2.Text = extractedText
        End If
    End Sub

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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim customText = TextBox2.Text
        Dim evaluatedResults As String = ProcessMathExpressions(customText)

        TextBox1.Text = TextBox1.Text & Environment.NewLine &
                        "Extracted Expression: " & customText & Environment.NewLine &
                        "Evaluated Results: " & evaluatedResults
        TextBox1.SelectionStart = TextBox1.Text.Length
        TextBox1.ScrollToCaret()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim pathToFile As String = Application.StartupPath & "templates\test data\"
        TestExtract(New Bitmap(pathToFile & "44 2.png"), "44 2", "((1 + 0) * (3 + 0)) * (55 - 44)")
        TestExtract(New Bitmap(pathToFile & "44.png"), "44", "(3 * 44) / (2 + 10)")
        TestExtract(New Bitmap(pathToFile & "33.png"), "33", "(6 + 33) / (96 - 93)")
        TestExtract(New Bitmap(pathToFile & "77.png"), "77", "((52 - 47) * (3 + 0)) + (23 + 77)")
        TestExtract(New Bitmap(pathToFile & "88.png"), "88", "((65 + 12) * (88 - 87)) + (35 + 0)")
        TestExtract(New Bitmap(pathToFile & "00.png"), "00", "((100 - 69) * (27 - 26)) + (1 * 73)")
    End Sub

    Private Sub TestExtract(img As Bitmap, filename As String, expected As String)
        Dim extractedText = TemplateMatching(img).Replace("-2", "2").Replace("=", "").Replace("?", "").Trim
        Dim evaluatedResults As String = ProcessMathExpressions(extractedText)

        Dim match = False
        If expected.Trim = extractedText.Trim Then
            match = True
        End If

        TextBox1.Text = TextBox1.Text & Environment.NewLine &
                        filename & Environment.NewLine &
                        "Expected: " & expected & Environment.NewLine &
                        "Extracted Expression: " & extractedText & Environment.NewLine &
                        "Match: " & match & Environment.NewLine &
                        "===================="
        TextBox1.SelectionStart = TextBox1.Text.Length
        TextBox1.ScrollToCaret()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click
        MessageBox.Show("Press overlay button to move or resize the overlay" & Environment.NewLine &
                        "arrow keys to move" & Environment.NewLine &
                        "+ and - for height" & Environment.NewLine &
                        "[ and ] for width" & Environment.NewLine &
                        "Press capture to extract and calculate" & Environment.NewLine & Environment.NewLine &
                        "Calc for incase the extracted expression is incorrect" & Environment.NewLine &
                        "You can manually input the expression in the textbox besides it" & Environment.NewLine,
                        "Help", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
End Class
