Public Enum InterpolationType
  Linear
  Logarithmic
End Enum

'See also: https://www.cmu.edu/biolphys/deserno/pdf/log_interpol.pdf
Public Function interpolate(x1 As Double, x2 As Double, y1 As Double, y2 As Double, x As Double, Optional tolerance As Double = 0, Optional interpType As InterpolationType = Linear) As Double
  'Is interpolation required?
  If withinTolerance(x, x1, tolerance) Then interpolate = y1: Exit Function
  If withinTolerance(x, x2, tolerance) Then interpolate = y2: Exit Function
  
  'Interpolate!
  If interpType = Linear Then
    'y=mx+c
    'm = (y2-y1)/(x2-x1)
    'y1 = m*x1+c --> c == y1 - m*x1
    'y<==m*x_new+c
    
    Dim m, c As Double
    m = (y2 - y1) / (x2 - x1)
    c = y1 - m * x1
    interpolate = m * x + c
  elseif interpType = Logarithmic then
    'Assume y' =mx' + c where t' = ln(t) then simple algorithm:
    'm = (y2'-y1')/(x2'-x1')
    'c = y1' - m*x1'
    'y' = m * x_new' + c
    'y = e^x'
    'And thus:
    'y = e^(linear_interpolate(ln(x1),ln(x2),ln(y1),ln(y2),ln(x),0,Linear))
    
    'ALTERNATIVE:
    interpolate = exp(interpolate(log(x1),log(x2),log(y1),log(y2),log(x),0,Linear))
  Else
    Err.Raise 0, "Main_SCF::interpolate", "Not implemented"
  End If
End Function