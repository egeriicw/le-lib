Attribute VB_Name = "RSubs"
Option Explicit





Public Sub Reg(X(), xr, xc, Y(), yr, yc, beta(), betar, betac, sec())
ReDim xt(xc, xr), xtx(xc, xc), xtxi(xc, xc), xtotal(xc, xr)
Dim xtr, xtc
Dim xtxr, xtxc
Dim xtotalr, xtotalc
Dim nv, i, j

'MATRIX RegRESSION ROUTINE

' THESE ARE MY SUB-PROGRAM TOOLS
' CALL Invert(A(), AR, AINV(), SEC())
' CALL Mult(A(), AR, AC, B(), BR, BC, AB(), ABR, ABC)
' CALL Trans(A(), AR, AC, AT(), ATR, ATC)
' WHERE  AR=NUM ROWS IN A()  AND  AC=NUM COLS IN A()

'CALCS X MEANS AND NORMALIZES X VARIABLES
nv = xc - 1
ReDim xt1(nv), xm(nv)
For j = 1 To nv
   xt1(j) = 0
Next
For i = 1 To xr
   For j = 1 To nv
     xt1(j) = xt1(j) + X(i, j + 1)
   Next
Next
For j = 1 To nv
   xm(j) = xt1(j) / xr
   If xm(j) = 0 Then xm(j) = 1
  For i = 1 To xr
      X(i, j + 1) = X(i, j + 1) / xm(j)
   Next
Next

Call Trans(X(), xr, xc, xt(), xtr, xtc)
Call Mult(xt(), xtr, xtc, X(), xr, xc, xtx(), xtxr, xtxc)
Call Invert(xtx(), xtxr, xtxi(), sec())
Call Mult(xtxi(), xtxr, xtxc, xt(), xtr, xtc, xtotal(), xtotalr, xtotalc)
Call Mult(xtotal(), xtotalr, xtotalc, Y(), yr, yc, beta(), betar, betac)

'RETURNS TO ORIGINAL X VALUES AND COEFFICIENTS
For i = 1 To xr
   For j = 1 To nv
      X(i, j + 1) = X(i, j + 1) * xm(j)
   Next
Next
For j = 1 To nv
   beta(j + 1, 1) = beta(j + 1, 1) / xm(j)
   sec(j + 1) = sec(j + 1) / (xm(j) ^ 2)
Next

End Sub


Public Sub Trans(X(), xr, xc, xt(), xtr, xtc)
Dim r, c

xtr = xc
xtc = xr

For r = 1 To xr
   For c = 1 To xc
      xt(c, r) = X(r, c)
   Next c
Next r

End Sub

Public Sub Invert(A(), s, B(), sec())
' THIS PROCEDURE WILL Invert UP TO AN 8 BY 8 MATRIX
' CHANGE THE DIMINSIONS IN THE "DIM INDEX()" STATEMENT FOR BIGGER MATRICES
ReDim Index(s, 3)
Dim det, numrows, numcols, errflag, i, j, k, l, big, hold, irow, icol, pivot
Dim t, m
numrows = s
numcols = s

On Error GoTo inves:

'TransFER A() TO B(), AND DO ALL FUTURE WORK ON B()
errflag = 0
For i = 1 To numrows
  For j = 1 To numcols
     B(i, j) = A(i, j)
  Next j
  Index(i, 3) = 0
Next i

det = 1
For i = 1 To numrows
   ' SEARCH FOR BIGGEST PIVOT ELEMENT
   big = 0
   For j = 1 To numcols
      If Index(j, 3) = 1 Then GoTo 5390
      For k = 1 To numcols
         If Index(k, 3) > 1 Then
            MsgBox "MATRIX IS SINGULAR", , "Error in Subroutine Invert"
            Stop
            Exit Sub
         End If
         If Index(k, 3) = 1 Then GoTo 5380
         If big >= Abs(B(j, k)) Then GoTo 5380
         
         irow = j
         icol = k
         big = Abs(B(j, k))
5380  Next k
5390  Next j
   Index(icol, 3) = Index(icol, 3) + 1
   Index(i, 1) = irow
   Index(i, 2) = icol
  
   'INTERCHANGE ROWS TO PUT PIVOT ON THE DIAGONAL
   If irow = icol Then GoTo 5580
   det = -det
   For l = 1 To numcols
      hold = B(irow, l)
      B(irow, l) = B(icol, l)
      B(icol, l) = hold
   Next l
  
5580 'DIVIDE PIVOT ROW BY PIVOT ELEMENT
    pivot = B(icol, icol)
    det = det * pivot
    B(icol, icol) = 1
    For l = 1 To numcols
       B(icol, l) = B(icol, l) / pivot
    Next l
  
    'REDUCE NON-PIVOT ROWS
    For l = 1 To numcols
       If l = icol Then GoTo 5810
       t = B(l, icol)
       B(l, icol) = 0
       For m = 1 To numcols
          B(l, m) = B(l, m) - B(icol, m) * t
       Next m
5810 Next l
Next i

'INTERCHANGE COLUMNS
For i = 1 To numcols
   l = numcols - i + 1
   If Index(l, 1) = Index(l, 2) Then GoTo 5960
   irow = Index(l, 1)
   icol = Index(l, 2)
   For k = 1 To numcols
      hold = B(k, irow)
      B(k, irow) = B(k, icol)
      B(k, irow) = hold
   Next k
5960 Next i

For k = 1 To numcols
   If Index(k, 3) <> 1 Then
      MsgBox "MATRIX IS SINGULAR", , "Error in Subroutine Invert"
      
      Exit Sub
   End If
Next k

For i = 1 To numcols
  sec(i) = B(i, i)
Next i

inves:

End Sub

Public Sub Inf(X(), xr, xc, Y(), yr, yc, beta(), betar, betac, sec(), rmse, cvrmse, r2, adjr2, sigma(), roe)
'CALULATES INFERENCE PARAMETERS
Dim yhatr, yhatc
Dim i, j, n, numcoefs, numvars, ytot, xtot, ymean
Dim acnum, acden, dwnum, dwden, prevres, autocor, dw, mse
Dim ssm, sse, ssy
Dim pnum, pdenom
ReDim xtot(xc), yhat(xr, 1), resx(xr)

'DEFINITIONS
n = xr
numvars = xc - 1
numcoefs = xc
ReDim xmean(numcoefs)
'CALCS X AND Y TOTALS AND MEANS
ytot = 0: ssy = 0
For j = 1 To numvars
   xtot(j) = 0
Next j
For i = 1 To n
   ytot = ytot + Y(i, 1)
   ssy = ssy + Y(i, 1) ^ 2
   For j = 1 To numvars
      xtot(j) = xtot(j) + X(i, j + 1)
   Next j
Next i
ymean = ytot / n
For i = 1 To numvars
   xmean(i) = xtot(i) / n
Next i

'CALCS SSE, MSE, RMSE, CV-RMSE AND R2
Call Mult(X(), xr, xc, beta(), betar, betac, yhat(), yhatr, yhatc)
ssm = 0: sse = 0
acnum = 0: acden = 0: dwnum = 0: dwden = 0
For i = 1 To n
   'fills resx() which corresponds to the x matrix
   resx(i) = Y(i, 1) - yhat(i, 1)
   ssm = ssm + (Y(i, 1) - ymean) ^ 2
   sse = sse + resx(i) ^ 2

   'calcs autocorr And durbinwatson
   Select Case i
   Case 1
       prevres = resx(i)
   Case Else
       acnum = acnum + prevres * resx(i)
       acden = acden + prevres ^ 2
       dwnum = dwnum + (resx(i) - prevres) ^ 2
       pnum = pnum + resx(i) * prevres
       pdenom = pdenom + resx(i) ^ 2
       prevres = resx(i)
   End Select
   dwden = dwden + resx(i) ^ 2
Next i
autocor = acnum / acden
dw = dwnum / dwden

If n - xc = 0 Then 'same number of data points as reg coef gives perfect fit
   mse = 0
Else
   mse = sse / (n - xc)
End If
rmse = mse ^ 0.5
cvrmse = rmse / ymean * 100
r2 = 1 - sse / ssm
If n - xc = 0 Then 'same number of data points as reg coef gives perfect fit
   adjr2 = 1
Else
   adjr2 = 1 - (n - 1) / (n - xc) * sse / ssm
End If
roe = pnum / pdenom

'CALCS THE STANDARD ERROR OF THE COEFFICIENTS
For i = 1 To numcoefs
   sigma(i) = rmse * sec(i) ^ 0.5
Next i

End Sub

Public Sub Mult(ma(), mar, mac, mb(), mbr, mbc, mc(), mcr, mcc)
Dim i, r, c

If mac <> mbr Then
     MsgBox "NUM X COL <> NUM Y ROWS", , "Error in Subroutine Mult"
   Exit Sub
End If

mcr = mar
mcc = mbc

For r = 1 To mcr
   For c = 1 To mcc
         mc(r, c) = 0
   Next c
Next r

For r = 1 To mcr
   For c = 1 To mcc
      For i = 1 To mac
        mc(r, c) = mc(r, c) + ma(r, i) * mb(i, c)
      Next i
   Next c
Next r

End Sub


