Public Class Form1st
    Dim rs As New FormResizer
    Private Sub Form1st_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        rs.FindAllControls(Me)
    End Sub
    Private Sub Form1st_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        rs.ResizeAllControls(Me)
    End Sub
    Public Class FormResizer
        '----------------------------------------------------------
        ' ControlInfo
        ' Structure of original state of all processed controls
        '----------------------------------------------------------
        Private Structure ControlInfo
            Public name As String
            Public parentName As String
            Public leftOffsetPercent As Double
            Public topOffsetPercent As Double
            Public heightPercent As Double
            Public originalHeight As Integer
            Public originalWidth As Integer
            Public widthPercent As Double
            Public originalFontSize As Single
        End Structure
        '-------------------------------------------------------------------------
        ' ctrlDict
        ' Dictionary of (control name, control info) for all processed controls
        '-------------------------------------------------------------------------
        Private ctrlDict As Dictionary(Of String, ControlInfo) = New Dictionary(Of String, ControlInfo)
        '----------------------------------------------------------------------------------------
        ' FindAllControls
        ' Recursive function to process all controls contained in the initially passed
        ' control container and store it in the Control dictionary
        '----------------------------------------------------------------------------------------
        Public Sub FindAllControls(thisCtrl As Control)
            '-- If the current control has a parent, store all original relative position
            '-- and size information in the dictionary.
            '-- Recursively call FindAllControls for each control contained in the
            '-- current Control
            For Each ctl As Control In thisCtrl.Controls
                Try
                    If Not IsNothing(ctl.Parent) Then
                        Dim parentHeight = ctl.Parent.Height
                        Dim parentWidth = ctl.Parent.Width
                        Dim c As New ControlInfo
                        c.name = ctl.Name
                        c.parentName = ctl.Parent.Name
                        c.topOffsetPercent = Convert.ToDouble(ctl.Top) / Convert.ToDouble(parentHeight)
                        c.leftOffsetPercent = Convert.ToDouble(ctl.Left) / Convert.ToDouble(parentWidth)
                        c.heightPercent = Convert.ToDouble(ctl.Height) / Convert.ToDouble(parentHeight)
                        c.widthPercent = Convert.ToDouble(ctl.Width) / Convert.ToDouble(parentWidth)
                        c.originalFontSize = ctl.Font.Size
                        c.originalHeight = ctl.Height
                        c.originalWidth = ctl.Width
                        ctrlDict.Add(c.name, c)
                    End If
                Catch ex As Exception
                    Debug.Print(ex.Message)
                End Try
                If ctl.Controls.Count > 0 Then
                    FindAllControls(ctl)
                End If
            Next '-- For Each
        End Sub
        '----------------------------------------------------------------------------------------
        ' ResizeAllControls
        ' Recursive function to resize and reposition all controls contained in the Control
        ' dictionary
        '----------------------------------------------------------------------------------------
        Public Sub ResizeAllControls(thisCtrl As Control)
            Dim fontRatioW As Single
            Dim fontRatioH As Single
            Dim fontRatio As Single
            Dim f As Font
            '-- Resize and reposition all controls in the passed control
            For Each ctl As Control In thisCtrl.Controls
                Try
                    If Not IsNothing(ctl.Parent) Then
                        Dim parentHeight = ctl.Parent.Height
                        Dim parentWidth = ctl.Parent.Width
                        Dim c As New ControlInfo
                        Dim ret As Boolean = False
                        Try
                            '-- Get the current control's info from the control info dictionary
                            ret = ctrlDict.TryGetValue(ctl.Name, c)
                            '-- If found, adjust the current control based on control relative
                            '-- size and position information stored in the dictionary
                            If (ret) Then
                                '-- Size
                                ctl.Width = Int(parentWidth * c.widthPercent)
                                ctl.Height = Int(parentHeight * c.heightPercent)
                                '-- Position
                                ctl.Top = Int(parentHeight * c.topOffsetPercent)
                                ctl.Left = Int(parentWidth * c.leftOffsetPercent)
                                '-- Font
                                f = ctl.Font
                                fontRatioW = ctl.Width / c.originalWidth
                                fontRatioH = ctl.Height / c.originalHeight
                                fontRatio = (fontRatioW +
                            fontRatioH) / 2 '-- average change in control Height and Width
                                ctl.Font = New Font(f.FontFamily,
                            c.originalFontSize * fontRatio, f.Style)
                            End If
                        Catch
                        End Try
                    End If
                Catch ex As Exception
                End Try
                '-- Recursive call for controls contained in the current control
                If ctl.Controls.Count > 0 Then
                    ResizeAllControls(ctl)
                End If
            Next '-- For Each
        End Sub
    End Class
    'End Class

    'Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Forms.Form2nd.Text = Now.ToString
        My.Forms.Form2nd.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim response As MsgBoxResult
        response = MsgBox("ΕΠΙΘΥΜΕΙΤΕ ΝΑ ΜΗ ΔΙΑΒΕΙΤΕ ΤΟ ΚΑΤΩΦΛΙ ΤΗΣ ΜΟΙΡΑΣ;", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirm")
        If response = MsgBoxResult.Yes Then
            Me.Dispose()
        ElseIf response = MsgBoxResult.No Then
            Exit Sub
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
