Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class OnOffLed
    Inherits Panel

    Public Enum LedState
        [On]
        Off
        OnSmall
        OnSmallYellow
        OffSmall
        OffSmallBlack
    End Enum

    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or
                 ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.ResizeRedraw Or
                 ControlStyles.UserPaint, True)
        UpdateStyles()
    End Sub

    Private _state As LedState = LedState.Off
    Public Property State As LedState
        Get
            Return _state
        End Get
        Set(value As LedState)
            _state = value
            Invalidate()
        End Set
    End Property

    Private _onText As String
    Public Property OnText As String
        Get
            Return _onText
        End Get
        Set(value As String)
            _onText = value
            Invalidate()
        End Set
    End Property

    Private _offText As String
    Public Property OffText As String
        Get
            Return _offText
        End Get
        Set(value As String)
            _offText = value
            Invalidate()
        End Set
    End Property

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim rec As New Rectangle(2, 2, Height - 5, Height - 5)
        Dim recText As New Rectangle(Height + 2, 1, Width - (Height - 2), Height)

        Dim G As Graphics = e.Graphics

        G.SmoothingMode = SmoothingMode.AntiAlias

        G.Clear(Parent.BackColor)

        ' large indicators
        If _state = LedState.On Then
            Dim cb As New ColorBlend With {
                .Colors = {Color.Orange, Color.Orange, Color.Orange},
                .Positions = {0, 0.5, 1}
            }

            Using lgb As New LinearGradientBrush(rec, Color.Empty, Color.Empty, 90.0F) With {.InterpolationColors = cb}
                G.FillRectangle(lgb, rec)
            End Using
        Else
            Dim cb As New ColorBlend With {
                .Colors = {Color.Black, Color.Black, Color.Black},
                .Positions = {0, 0.5, 1}
            }
            Using lgb As New LinearGradientBrush(rec, Color.Empty, Color.Empty, 90.0F) With {.InterpolationColors = cb}
                G.FillRectangle(lgb, rec)
            End Using
        End If

        ' small indicators
        If _state = LedState.OnSmall Then
            Dim cb As New ColorBlend With {
                .Colors = {Color.LightGreen, Color.LightGreen, Color.LightGreen},
                .Positions = {0, 0.5, 1}
            }

            Using lgb As New LinearGradientBrush(rec, Color.Empty, Color.Empty, 90.0F) With {.InterpolationColors = cb}
                G.FillRectangle(lgb, rec)
            End Using
        End If
        If _state = LedState.OnSmallYellow Then
            Dim cb As New ColorBlend With {
                .Colors = {Color.Yellow, Color.Yellow, Color.Yellow},
                .Positions = {0, 0.5, 1}
            }

            Using lgb As New LinearGradientBrush(rec, Color.Empty, Color.Empty, 90.0F) With {.InterpolationColors = cb}
                G.FillRectangle(lgb, rec)
            End Using
        End If
        If _state = LedState.OffSmall Then
            Dim cb As New ColorBlend With {
                    .Colors = {Color.Red, Color.Red, Color.Red},
                    .Positions = {0, 0.5, 1}
                }
            Using lgb As New LinearGradientBrush(rec, Color.Empty, Color.Empty, 90.0F) With {.InterpolationColors = cb}
                G.FillRectangle(lgb, rec)
            End Using
        End If
        If _state = LedState.OffSmallBlack Then
            Dim cb As New ColorBlend With {
                .Colors = {Color.Black, Color.Black, Color.Black},
                .Positions = {0, 0.5, 1}
            }

            Using lgb As New LinearGradientBrush(rec, Color.Empty, Color.Empty, 90.0F) With {.InterpolationColors = cb}
                G.FillRectangle(lgb, rec)
            End Using
        End If

        G.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

        Using br As New SolidBrush(ForeColor)
            Using sf As New StringFormat With {.Alignment = StringAlignment.Near, .LineAlignment = StringAlignment.Center}
                G.DrawString(If(_state = LedState.On, _onText, _offText), Font, br, recText, sf)
            End Using
        End Using

        ' large indicators only
        If _state = LedState.On Or _state = LedState.Off Then
            ' draw outline around every filled square - IanJ
            Dim myPen As New Pen(Color.Black, 1) 'to set width of pen
            e.Graphics.DrawRectangle(myPen, 3, 3, Width - 6, Height - 6)  ' this is a bit of a hack
        End If
    End Sub
End Class
