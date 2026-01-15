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

    ' =========================================================================
    '  METROLOGY / CALIBRATION HELPERS

    ' Returns only real roots and sorts ascending.
    ' Also clamps tiny -0 to 0 for nicer display.
    Public Function RealRootsSorted(roots As Global.System.Numerics.Complex(),
                                   Optional tol As Double = 0.0000000001) As Double()
        Dim rr = RealRoots(roots, tol)
        Array.Sort(rr)
        For i As Integer = 0 To rr.Length - 1
            If Math.Abs(rr(i)) < tol Then rr(i) = 0R
        Next
        Return rr
    End Function

    ' Two-point linear calibration.
    ' Returns {gain, offset} such that: true = gain * measured + offset
    Public Function TwoPointCal(meas1 As Double, true1 As Double,
                               meas2 As Double, true2 As Double) As Double()
        If meas2 = meas1 Then Throw New ArgumentException("meas2 must differ from meas1")
        Dim gain As Double = (true2 - true1) / (meas2 - meas1)
        Dim offset As Double = true1 - gain * meas1
        Return {gain, offset}
    End Function

    ' Apply linear calibration: true = gain*measured + offset
    Public Function ApplyCal(measured As Double, gain As Double, offset As Double) As Double
        Return gain * measured + offset
    End Function

    ' Parts-per-million error relative to nominal: (measured-nominal)/nominal * 1e6
    Public Function PpmError(measured As Double, nominal As Double) As Double
        If nominal = 0R Then Throw New ArgumentException("nominal must be non-zero")
        Return (measured - nominal) / nominal * 1000000.0R
    End Function

    ' Least-squares straight line fit: y = m*x + b
    ' Returns {m, b, r2}
    Public Function LineFit(xs() As Double, ys() As Double) As Double()
        If xs Is Nothing OrElse ys Is Nothing Then Throw New ArgumentException("xs/ys must not be Nothing")
        If xs.Length <> ys.Length OrElse xs.Length < 2 Then Throw New ArgumentException("xs/ys must be same length >= 2")

        Dim n As Integer = xs.Length
        Dim sx As Double = 0R, sy As Double = 0R, sxx As Double = 0R, sxy As Double = 0R
        For i As Integer = 0 To n - 1
            Dim x = xs(i) : Dim y = ys(i)
            sx += x : sy += y
            sxx += x * x
            sxy += x * y
        Next

        Dim denom As Double = n * sxx - sx * sx
        If denom = 0R Then Throw New ArgumentException("LineFit: singular (all x identical?)")

        Dim m As Double = (n * sxy - sx * sy) / denom
        Dim b As Double = (sy - m * sx) / n

        ' r^2
        Dim ymean As Double = sy / n
        Dim ssTot As Double = 0R
        Dim ssRes As Double = 0R
        For i As Integer = 0 To n - 1
            Dim yi = ys(i)
            Dim fi = m * xs(i) + b
            ssTot += (yi - ymean) * (yi - ymean)
            ssRes += (yi - fi) * (yi - fi)
        Next
        Dim r2 As Double = If(ssTot = 0R, 1.0R, 1.0R - (ssRes / ssTot))

        Return {m, b, r2}
    End Function

    ' Polynomial derivative.
    ' coeffs = {a0,a1,a2,...} => d/dx = {a1, 2*a2, 3*a3, ...}
    Public Function PolyDer(coeffs() As Double) As Double()
        If coeffs Is Nothing OrElse coeffs.Length <= 1 Then Return Array.Empty(Of Double)()
        Dim out(coeffs.Length - 2) As Double
        For i As Integer = 1 To coeffs.Length - 1
            out(i - 1) = coeffs(i) * i
        Next
        Return out
    End Function

    ' Polynomial integral.
    ' coeffs = {a0,a1,a2,...} => ∫ = {c0, a0/1, a1/2, a2/3, ...}
    Public Function PolyInt(coeffs() As Double, Optional c0 As Double = 0R) As Double()
        If coeffs Is Nothing OrElse coeffs.Length = 0 Then Return {c0}
        Dim out(coeffs.Length) As Double
        out(0) = c0
        For i As Integer = 0 To coeffs.Length - 1
            out(i + 1) = coeffs(i) / (i + 1)
        Next
        Return out
    End Function

    ' Least-squares polynomial fit (metrology calibration curves).
    ' Returns coefficients {a0,a1,...,adeg} for: y = a0 + a1*x + ... + adeg*x^deg
    Public Function PolyFit(xs() As Double, ys() As Double, degree As Integer) As Double()
        If xs Is Nothing OrElse ys Is Nothing Then Throw New ArgumentException("xs/ys must not be Nothing")
        If xs.Length <> ys.Length OrElse xs.Length = 0 Then Throw New ArgumentException("xs/ys must be same non-zero length")
        If degree < 0 Then Throw New ArgumentException("degree must be >= 0")
        If xs.Length < degree + 1 Then Throw New ArgumentException("need at least degree+1 points")

        Dim n As Integer = xs.Length
        Dim m As Integer = degree + 1

        ' Build normal equations: (V^T V) a = V^T y
        Dim ata(m - 1, m - 1) As Double
        Dim aty(m - 1) As Double

        For i As Integer = 0 To n - 1
            Dim x As Double = xs(i)
            Dim y As Double = ys(i)

            ' powers(0)=1, powers(1)=x, ...
            Dim powers(m - 1) As Double
            powers(0) = 1.0R
            For p As Integer = 1 To m - 1
                powers(p) = powers(p - 1) * x
            Next

            For row As Integer = 0 To m - 1
                aty(row) += powers(row) * y
                For col As Integer = 0 To m - 1
                    ata(row, col) += powers(row) * powers(col)
                Next
            Next
        Next

        Return SolveLinearSystemGaussian(ata, aty)
    End Function

    ' Gaussian elimination with partial pivoting.
    Private Function SolveLinearSystemGaussian(a(,) As Double, b() As Double) As Double()
        Dim n As Integer = b.Length
        Dim Awork(n - 1, n - 1) As Double
        Dim x(n - 1) As Double
        Dim bwork(n - 1) As Double

        Array.Copy(b, bwork, n)
        For i As Integer = 0 To n - 1
            For j As Integer = 0 To n - 1
                Awork(i, j) = a(i, j)
            Next
        Next

        For k As Integer = 0 To n - 1
            ' pivot
            Dim piv As Integer = k
            Dim maxAbs As Double = Math.Abs(Awork(k, k))
            For i As Integer = k + 1 To n - 1
                Dim v = Math.Abs(Awork(i, k))
                If v > maxAbs Then
                    maxAbs = v
                    piv = i
                End If
            Next
            If maxAbs = 0R Then Throw New ArgumentException("Singular matrix in PolyFit/LineFit")

            If piv <> k Then
                For j As Integer = k To n - 1
                    Dim tmp = Awork(k, j)
                    Awork(k, j) = Awork(piv, j)
                    Awork(piv, j) = tmp
                Next
                Dim tb = bwork(k) : bwork(k) = bwork(piv) : bwork(piv) = tb
            End If

            ' eliminate
            Dim akk As Double = Awork(k, k)
            For i As Integer = k + 1 To n - 1
                Dim f As Double = Awork(i, k) / akk
                Awork(i, k) = 0R
                For j As Integer = k + 1 To n - 1
                    Awork(i, j) -= f * Awork(k, j)
                Next
                bwork(i) -= f * bwork(k)
            Next
        Next

        ' back-substitute
        For i As Integer = n - 1 To 0 Step -1
            Dim s As Double = bwork(i)
            For j As Integer = i + 1 To n - 1
                s -= Awork(i, j) * x(j)
            Next
            x(i) = s / Awork(i, i)
        Next

        Return x
    End Function
End Module