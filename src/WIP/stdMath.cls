VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum InterpolationType
  Linear
  Logarithmic  
  Hyperbolic   'https://www.engineerexcel.com/hyperbolic-curve-fitting-excel/  y=mx/(k+x) --> (1/y) = k/m (1/x) + 1/m
End Enum

Private Type InterpolationStruct
  InterpType as InterpolationType
  LinearM as double
  LinearC as double
End Type

'See also: https://www.cmu.edu/biolphys/deserno/pdf/log_interpol.pdf
Public Function interpolate(x1 As Double, x2 As Double, y1 As Double, y2 As Double, x As Double, Optional tolerance As Double = 0, Optional interpType As Long = Linear) As Double
  'Is interpolation required?
  If withinTolerance(x, x1, tolerance) Then interpolate = y1: Exit Function
  If withinTolerance(x, x2, tolerance) Then interpolate = y2: Exit Function
  
  'Interpolate!
  If interpType = Linear Then
    'y=mx+c
    'm = (y2-y1)/(x2-x1)
    'y1 = m*x1+c --> c == y1 - m*x1
    'y<==m*x_new+c
    
    Dim ip as InterpolationStruct
    ip = GetLinearInterpolateStruct(x1,x2,y1,y2)
    interpolate = ip.LinearM * x + ip.LinearC
  elseif interpType = Logarithmic then
    'Assume y' =mx' + c where t' = ln(t) then simple algorithm:
    'm = (y2'-y1')/(x2'-x1')
    'c = y1' - m*x1'
    'y' = m * x_new' + c
    'y = e^x'
    'And thus:
    'y = e^(linear_interpolate(ln(x1),ln(x2),ln(y1),ln(y2),ln(x),0,Linear))
    
    'ALTERNATIVE:
    interpolate = exp(interpolate(log(x1),log(x2),log(y1),log(y2),log(x),tolerance,Linear))
  elseif interpType = Hyperbolic then
    'Assume y = (mx)/(k+x)
    'This formula can be re-arranged as follows:
    '    (1/y) = (k/m) * (1/x) + 1/m
    '==> Y = MX + C
    '   Linearised function can be solved for m and c
    Dim ip as InterpolationStruct
    ip = GetLinearInterpolateStruct(1/x1,1/x2,1/y1,1/y2)
    
    'Get M and C   for    Y = MX +C
    Dim M as double, C as double
    M = ip.LinearM
    C = ip.LinearC

    'Get m and k
    Dim m_ as double:  m_ = 1/C   'C = 1/m ==> m = 1/C
    Dim k_ as double   k_ = M*m_  'M = k/m ==> k = M*m

    'Finalise interpolation
    'y = (mx)/(k+x)
    interpolate = (m_ * x) / (k_ + x)
  Else
    Err.Raise 0, "Main_SCF::interpolate", "Not implemented"
  End If
End Function

Private Function GetLinearInterpolateStruct(x1 as double, x2 as double, y1 as double, y2 as double) as InterpolationStruct
  Dim x as InterpolationStruct
  x.InterpType = Linear
  x.LinearM = (y2 - y1) / (x2 - x1)
  x.LinearC = y1 - m * x1
  GetLinearInterpolateStruct = x
End Function


'fit a curve to data - fit(arr2D() as double)
Public Function fit(arr2d() as double, iType as Long) as InterpolationStruct
  '
End Function
