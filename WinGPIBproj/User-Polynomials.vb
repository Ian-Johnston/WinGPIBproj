Option Strict On
Option Explicit On

Imports System

' If the project has ANY "System" namespace shadowing, keep using Global.System.Numerics.Complex below.
' Otherwise change Global.System.Numerics.Complex -> System.Numerics.Complex.
Imports System.Numerics

Public Module PolynomialSolvers

    ' =========================================================================
    '  POLYNOMIAL EVALUATION (recommended for calibration curves)

    ' Evaluates y = a0 + a1*x + a2*x^2 + ... using Horner's method (fast/stable).
    ' coeffs = {a0, a1, a2, ...}  (constant term first)
    Public Function PolyEval(x As Double, ParamArray coeffs() As Double) As Double
        If coeffs Is Nothing OrElse coeffs.Length = 0 Then Return 0R
        Dim y As Double = 0R
        For i As Integer = coeffs.Length - 1 To 0 Step -1
            y = y * x + coeffs(i)
        Next
        Return y
    End Function

    ' =========================================================================
    '  COMPLEX HELPERS (for solvers)

    Private Function Cbrt(z As Global.System.Numerics.Complex) As Global.System.Numerics.Complex
        ' Principal complex cube root
        If z = Global.System.Numerics.Complex.Zero Then Return Global.System.Numerics.Complex.Zero

        Dim r As Double = z.Magnitude
        Dim theta As Double = Math.Atan2(z.Imaginary, z.Real)

        Dim rr As Double = Math.Pow(r, 1.0R / 3.0R)
        Dim tt As Double = theta / 3.0R

        Return New Global.System.Numerics.Complex(rr * Math.Cos(tt), rr * Math.Sin(tt))
    End Function

    Private Function NormalizeLeading(a As Double, ParamArray coeffs() As Double) As Double()
        Dim out(coeffs.Length - 1) As Double
        For i As Integer = 0 To coeffs.Length - 1
            out(i) = coeffs(i) / a
        Next
        Return out
    End Function

    ' =========================================================================
    '  QUADRATIC SOLVER (Complex roots)

    ' Solves: a*x^2 + b*x + c = 0
    Public Function SolveQuadratic(a As Double, b As Double, c As Double) As Global.System.Numerics.Complex()
        If a = 0R Then Throw New ArgumentException("a must be non-zero for quadratic")

        Dim disc As Global.System.Numerics.Complex = New Global.System.Numerics.Complex(b * b - 4.0R * a * c, 0R)
        Dim sqrtDisc As Global.System.Numerics.Complex = Global.System.Numerics.Complex.Sqrt(disc)
        Dim denom As Double = 2.0R * a

        Return {
            (-b + sqrtDisc) / denom,
            (-b - sqrtDisc) / denom
        }
    End Function

    Private Function SolveQuadraticComplex(a As Double,
                                          b As Global.System.Numerics.Complex,
                                          c As Global.System.Numerics.Complex) As Global.System.Numerics.Complex()
        Dim disc As Global.System.Numerics.Complex = b * b - 4.0R * a * c
        Dim sqrtDisc As Global.System.Numerics.Complex = Global.System.Numerics.Complex.Sqrt(disc)
        Dim denom As Double = 2.0R * a

        Return {
            (-b + sqrtDisc) / denom,
            (-b - sqrtDisc) / denom
        }
    End Function

    ' =========================================================================
    '  CUBIC SOLVER (Cardano, Complex roots)

    ' Solves: a*x^3 + b*x^2 + c*x + d = 0
    Public Function SolveCubic(a As Double, b As Double, c As Double, d As Double) As Global.System.Numerics.Complex()
        If a = 0R Then Throw New ArgumentException("a must be non-zero for cubic")

        ' Normalize to monic: x^3 + ba*x^2 + ca*x + da = 0
        Dim n = NormalizeLeading(a, b, c, d)
        Dim ba As Double = n(0)
        Dim ca As Double = n(1)
        Dim da As Double = n(2)

        ' Depress: x = y - ba/3
        Dim p As Double = ca - (ba * ba) / 3.0R
        Dim q As Double = (2.0R * ba * ba * ba) / 27.0R - (ba * ca) / 3.0R + da

        Dim halfQ As Global.System.Numerics.Complex = New Global.System.Numerics.Complex(q / 2.0R, 0R)
        Dim disc As Global.System.Numerics.Complex =
            halfQ * halfQ + New Global.System.Numerics.Complex((p / 3.0R) * (p / 3.0R) * (p / 3.0R), 0R)

        Dim u As Global.System.Numerics.Complex = Cbrt(-halfQ + Global.System.Numerics.Complex.Sqrt(disc))
        Dim v As Global.System.Numerics.Complex = Cbrt(-halfQ - Global.System.Numerics.Complex.Sqrt(disc))

        ' Cube roots of unity
        Dim w As New Global.System.Numerics.Complex(-0.5R, Math.Sqrt(3.0R) / 2.0R)
        Dim w2 As Global.System.Numerics.Complex = w * w

        Dim y1 As Global.System.Numerics.Complex = u + v
        Dim y2 As Global.System.Numerics.Complex = w * u + w2 * v
        Dim y3 As Global.System.Numerics.Complex = w2 * u + w * v

        Dim shift As Double = ba / 3.0R
        Return {y1 - shift, y2 - shift, y3 - shift}
    End Function

    ' =========================================================================
    '  QUARTIC SOLVER (Ferrari, Complex roots)

    ' Solves: a*x^4 + b*x^3 + c*x^2 + d*x + e = 0
    Public Function SolveQuartic(a As Double, b As Double, c As Double, d As Double, e As Double) As Global.System.Numerics.Complex()
        If a = 0R Then Throw New ArgumentException("a must be non-zero for quartic")

        ' Normalize to monic: x^4 + Bb*x^3 + Cc*x^2 + Dd*x + Ee = 0
        Dim n = NormalizeLeading(a, b, c, d, e)
        Dim Bb As Double = n(0)
        Dim Cc As Double = n(1)
        Dim Dd As Double = n(2)
        Dim Ee As Double = n(3)

        ' Depress: x = y - Bb/4
        Dim Bb2 As Double = Bb * Bb
        Dim p As Double = Cc - 3.0R * Bb2 / 8.0R
        Dim q As Double = Dd + (Bb2 * Bb) / 8.0R - (Bb * Cc) / 2.0R
        Dim r As Double = Ee - 3.0R * Bb2 * Bb2 / 256.0R + (Bb2 * Cc) / 16.0R - (Bb * Dd) / 4.0R

        ' Resolvent cubic:
        ' z^3 - (p/2) z^2 - r z + (r p/2 - q^2/8) = 0
        Dim rcA As Double = 1.0R
        Dim rcB As Double = -p / 2.0R
        Dim rcC As Double = -r
        Dim rcD As Double = r * p / 2.0R - (q * q) / 8.0R

        Dim zRoots = SolveCubic(rcA, rcB, rcC, rcD)

        ' Pick one root (largest real part tends to be a bit more stable)
        Dim z As Global.System.Numerics.Complex = zRoots(0)
        For i As Integer = 1 To 2
            If zRoots(i).Real > z.Real Then z = zRoots(i)
        Next

        Dim U As Global.System.Numerics.Complex = Global.System.Numerics.Complex.Sqrt(2.0R * z - p)

        Dim V As Global.System.Numerics.Complex
        If U = Global.System.Numerics.Complex.Zero Then
            V = Global.System.Numerics.Complex.Zero
        Else
            V = New Global.System.Numerics.Complex(-q, 0R) / (2.0R * U)
        End If

        ' Two quadratics:
        ' y^2 + U*y + (z - V) = 0
        ' y^2 - U*y + (z + V) = 0
        Dim yRoots1 = SolveQuadraticComplex(1.0R, U, z - V)
        Dim yRoots2 = SolveQuadraticComplex(1.0R, -U, z + V)

        Dim shift As Double = Bb / 4.0R

        Return {
            yRoots1(0) - shift,
            yRoots1(1) - shift,
            yRoots2(0) - shift,
            yRoots2(1) - shift
        }
    End Function

    ' =========================================================================
    '  OPTIONAL: REAL-ROOT FILTER (handy for UI/calibration)

    ' Returns only roots where |Imag| <= tol
    Public Function RealRoots(roots As Global.System.Numerics.Complex(), Optional tol As Double = 0.0000000001) As Double()
        If roots Is Nothing Then Return Array.Empty(Of Double)()

        Dim lst As New List(Of Double)()
        For Each z In roots
            If Math.Abs(z.Imaginary) <= tol Then
                lst.Add(z.Real)
            End If
        Next
        Return lst.ToArray()
    End Function

End Module
