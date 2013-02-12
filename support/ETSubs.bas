Attribute VB_Name = "ETSubs"


Public Sub FillD(filename$)
Dim l$, delim, charnum
Dim i, j, scharnum, fnum, echarnum, flength, row, col

'sets default no data flag
ndflag = -99

'finds delimiter type, numflds, numrecs
numrecs = 0
numflds = 0
Open filename$ For Input As #1
While Not EOF(1)
   'reads a line
   Line Input #1, l$
   l$ = Trim$(l$)
   If (Len(l$) > 0) And (Left(l$, 1) = "1" Or Left(l$, 1) = "2" Or Left(l$, 1) = "3" Or Left(l$, 1) = "4" Or Left(l$, 1) = "5" Or Left(l$, 1) = "6" Or Left(l$, 1) = "7" Or Left(l$, 1) = "8" Or Left(l$, 1) = "9" Or Left(l$, 1) = "0") Then
      numrecs = numrecs + 1
   
      If numrecs = 1 Then
         
         'determines if comma, tab (chr(9)), or space (chr(32)) delimited
         delim = ""
         For charnum = 1 To Len(l$)
            If Mid$(l$, charnum, 1) = "," Then
               delim = ","
               Exit For
            End If
         Next charnum
         If delim <> "," Then
            For charnum = 1 To Len(l$)
               If Mid$(l$, charnum, 1) = Chr(9) Then
                  delim = Chr(9)
                  Exit For
               End If
            Next charnum
         End If
         If delim = "" Then delim = Chr(32)
   
         'determines numflds and dims field()
         numflds = 0
         If delim = "," Or delim = Chr(9) Then
            For charnum = 1 To Len(l$)
               If Mid$(l$, charnum, 1) = delim Or charnum = Len(l$) Then
                  numflds = numflds + 1
               End If
            Next charnum
         Else 'delim = " "
            For charnum = 1 To Len(l$)
               If Mid$(l$, charnum, 1) = delim Then
                  If Mid$(l$, charnum - 1, 1) <> delim Then numflds = numflds + 1
               ElseIf charnum = Len(l$) Then
                  numflds = numflds + 1
               End If
            Next charnum
         End If
      End If
   End If
Wend
Close #1
ReDim d(numrecs, numflds)


'fills d(row,col) if space, tab or comma delim
row = 0
Open filename$ For Input As #1
While Not EOF(1)
   'reads a line
   Line Input #1, l$
   l$ = Trim$(l$)
   If (Len(l$) > 0) And (Left(l$, 1) = "1" Or Left(l$, 1) = "2" Or Left(l$, 1) = "3" Or Left(l$, 1) = "4" Or Left(l$, 1) = "5" Or Left(l$, 1) = "6" Or Left(l$, 1) = "7" Or Left(l$, 1) = "8" Or Left(l$, 1) = "9" Or Left(l$, 1) = "0") Then
      row = row + 1
      'parses the line into fields
      scharnum = 1
      fnum = 0
      For charnum = 1 To Len(l$)
         If Mid$(l$, charnum, 1) = delim Then
            If Mid$(l$, charnum - 1, 1) <> delim Then
               fnum = fnum + 1
               If fnum > numflds Then
                  msg$ = "Incorrect number of columns in line " + Str$(row) + " of " + UCase$(filename$) + "."
                  MsgBox msg$, , "Error"
                  Close #1
                  Screen.MousePointer = 0
                  Exit Sub
               End If
               echarnum = charnum - 1
               flength = echarnum - scharnum + 1
               d(row, fnum) = Val(Trim(Mid$(l$, scharnum, flength)))
               scharnum = echarnum + 2
            End If
         ElseIf charnum = Len(l$) Then
            fnum = fnum + 1
            echarnum = charnum
            flength = echarnum - scharnum + 1
            d(row, fnum) = Val(Trim(Mid$(l$, scharnum, flength)))
         End If
      Next charnum
   End If
   Call UpdatePerDone(row, numrecs * 2, 1)
Wend

Close #1

End Sub



Public Sub Merge()
Dim i, j, k
Dim startrec, startdate, enddate, dydate, n, tot
Dim regnum, totrecs, increment 'for perdone
On Error GoTo mergece

'this subroutine integrates e(mo,dy,yr,elec,dem,ng, elecpp, ngpp, p1, p2) and w(mo,dy,yr,T) into d(mo,dy,yr,elec,dem,ng, elecpp, ngpp, T)

'show perdone
regnum = 0
totregs = wrecs
increment = 10
PerDone.Caption = "Merging Energy and Weather Files..."
PerDone.Show 0
DoEvents

'creates d() and fills projected energy use fields with nodata flags
ReDim d(0 To erecs, 28)
For i = 1 To erecs
   If numindvars = 1 Then
      For j = 9 To 28
         d(i, j) = -99
      Next j
   ElseIf numindvars = 2 Then
      For j = 10 To 28
         d(i, j) = -99
      Next j
   Else
      For j = 11 To 28
         d(i, j) = -99
      Next j
   End If
   
Next i

'fills d(0, ) with dates and nodata flags
For i = 1 To 10
   d(0, i) = e(0, i)
Next i

'integrates temp data from w() with energy data from e() and puts results in d()
startrec = 1
For i = 1 To erecs
       
   'fills d() with first 10 fields of e()
   For k = 1 To 10
      d(i, k) = e(i, k)
   Next k
   d(i, 11) = -99

   'finds start and end date for each period in monthly file
   startdate = CDate(Format(e(i - 1, 1)) + "/" + Format(e(i - 1, 2)) + "/" + Format(e(i - 1, 3)))
   enddate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
    
   'finds average temperature in energy period
   n = 0: ttot = 0
   For j = startrec To wrecs
      dydate = CDate(Format(w(j, 1)) + "/" + Format(w(j, 2)) + "/" + Format(w(j, 3)))
      
      If dydate > startdate And dydate <= enddate And w(j, 4) <> -99 Then
         'finds average
         n = n + 1
         ttot = ttot + w(j, 4)
         d(i, 11) = ttot / n
         
         'set new startrec
         startrec = j
      End If
            
      'updates perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, increment)
      
      If dydate > enddate Then Exit For
   Next j

Next i
   
'sets global params
numrecs = erecs
numflds = 14 '8
mofld = 1
dyfld = 2
yrfld = 3
'hrfld = 4
timeint = 3

'sets global variable names
ReDim vn$(28)
vn$(1) = "Mo"
vn$(2) = "Dy"
vn$(3) = "Yr"
vn$(4) = "Elec Use (kWh/mo)"
vn$(5) = "Elec Demand (kW)"
vn$(6) = "Fuel Use (units/mo)"
vn$(7) = "ElecPrePst (kWh/mo)"
vn$(8) = "FuelPrePst (units/mo)"
vn$(9) = "Prod/Occ 1 (units/mo)"
vn$(10) = "Prod/Occ 2 (units/mo)"
vn$(11) = "Avg Temp (F)"
vn$(12) = "Elec Use if no retrofit (kWh/mo)"
vn$(13) = "Elec Demand if no retrofit (kW)"
vn$(14) = "Fuel Use if no retrofit (units/mo)"
vn$(15) = "Annual Consumption (units/year)"
vn$(16) = "Normal Annual Consumption (units/year)"
vn$(17) = "Tbalance (F)" '"XCP1"
vn$(18) = "XCP2"
vn$(19) = "Weather Independent Use"  '"YCP"
vn$(20) = "Slope = UA/Eff" '"LS"
vn$(21) = "Slope = UA/Eff" '"RS"
vn$(22) = "IV1"
vn$(23) = "IV2"
vn$(24) = "N"
vn$(25) = "Ymean"
vn$(26) = "R2"
vn$(27) = "RMSE"
vn$(28) = "CVRMSE"


'resets paths to files
'tnp$ = tnpmo$
'tfn$ = tfnmo$

'clears arrays to save memory
'ReDim w(1, 1), e(1, 1)

mergece:
Screen.MousePointer = 0
Unload PerDone
PerDone.Caption = "Processing Data..."
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume mergees
End If
mergees:

End Sub


Public Sub UpdatePerDone(i, lasti, inc)

'logic of subroutine
'for i = 1 to lasti
   'update perdone when i is a multiple of inc
'next i

'update perdone
If Int(i / inc) * inc = i And lasti > 0 Then
   PerDone.Picture1.Line (0, 0)-(i / lasti * PerDone.Picture1.Width, PerDone.Picture1.Height), BLUE, BF
End If

End Sub

Public Sub FourP(xcp, b1, b2, b3, cvrmse, rmse, r2)
ReDim A(numrecs, 3), rmsei(100)
Dim numints As Integer, itration As Integer
Dim i As Integer, j As Integer, n1 As Integer, n2 As Integer
Dim numpasses As Integer, count As Integer
Dim totregs, numregs, xmin4p
Dim numxvars, xr, xc, inc, cp, IndVar, rmsemin, bestcp, regnum, t, predint
On Error GoTo cpce

'show perdone
Screen.MousePointer = 11
PerDone.Caption = "Processing Data..."
PerDone.Show 0
DoEvents

'call SetAll to fill x(),y(), min, max etc
'Call SetAll
'ReDim IndexNumO(n) As Integer, WDWEO(n) As Integer, XO(n), YO(n)

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo cpce:
End If

modtype = "4P"

'sets xrows and x columns
xr = n
xc = 2

'dimension arrays
'ReDim x(numrecs, numxvars + 1), y(numrecs, numxvars + 1)
numxvars = xc
ReDim beta(numxvars + 1, 1), sec(numxvars + 1), sigma(numxvars + 1)

'Calcs initial size of search grid
numints = 10
inc = (xmax - xmin) / numints

'find total regressions to be performed for perdone
totregs = 2 * (numints - 1) + 1 'for fine search

'FINDS CP BY USING COURSE AND THEN FINER SEARCHES, ITRATION = NUM OF SEARCHES
For itration = 1 To 2
   
   'set xmin4p and inc depending on course or fine grid search
   If itration = 1 Then
      xmin4p = xmin
   Else '(itration = 2 for fine grid...
      xmin4p = bestcp - inc
      inc = 2 * inc / numints
   End If

   For j = 1 To numints - 1
      rmsei(j) = 0
   Next j
   
   For i = 1 To numints - 1   'FOR EACH INCREMENT....
      'INITIALIZE COUNTERS AND ARRAYS
      n1 = 0: n2 = 0

      'FILL A()
      cp = xmin4p + i * inc
      For j = 1 To n
         If x(j, 2) <= cp Then
            n1 = n1 + 1
            IndVar = 0
         Else
            n2 = n2 + 1
            IndVar = 1
         End If
         A(j, 1) = 1
         A(j, 2) = x(j, 2)
         A(j, 3) = IndVar * (x(j, 2) - cp)
      Next j

      'prints error if n = 0
      If n = 0 Then
         MsgBox "No data is available to model.", , "Error"
         Screen.MousePointer = 0
         GoTo cpce:
      End If

      'Call Reg(A(), n, 3, y(), n, 1, beta(), betar, betac, sec())
      'Call Inf(A(), n, 3, y(), n, 1, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), P)
      Call Reg(A(), n, numxvars + 1, y(), n, 1, beta(), numxvars + 1, 1, sec())
      Call Inf(A(), n, numxvars + 1, y(), n, 1, beta(), numxvars + 1, 1, sec(), rmse, cvrmse, r2, adjr2, sigma(), roe)

      'SELECTS CP AT MINIMUM RMSE
      rmsei(i) = rmse
      If i = 1 Then
         rmsemin = rmsei(1)
         bestcp = cp
      End If
      If rmsei(i) < rmsemin Then
         rmsemin = rmsei(i)
         bestcp = cp
      End If
      
      'update perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, 1)
   
   Next i
Next itration

'USES BEST CP TO FIND FINAL STATS
n1 = 0: n2 = 0
cp = bestcp
For j = 1 To n
   If x(j, 2) <= cp Then
      n1 = n1 + 1
      IndVar = 0
   Else
      n2 = n2 + 1
      IndVar = 1
   End If
   A(j, 1) = 1
   A(j, 2) = x(j, 2)
   A(j, 3) = IndVar * (x(j, 2) - cp)
Next j
'Call Reg(A(), n, 3, y(), n, 1, beta(), betar, betac, sec())
'Call Inf(A(), n, 3, y(), n, 1, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), P)
Call Reg(A(), n, numxvars + 1, y(), n, 1, beta(), numxvars + 1, 1, sec())
Call Inf(A(), n, numxvars + 1, y(), n, 1, beta(), numxvars + 1, 1, sec(), rmse, cvrmse, r2, adjr2, sigma(), roe)

'xcp = bestcp
'ycp = beta(1, 1) + beta(2, 1) * bestcp
'ls = beta(2, 1)
'rs = beta(2, 1) + beta(3, 1)
xcp = bestcp
b1 = beta(1, 1)
b2 = beta(2, 1)
b3 = beta(3, 1)

'calcs prediction interval
t = 1.96 + 2.7 / (n - 3)
predint = t * rmse * (1 + 1 / n) ^ 0.5

cpce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   MsgBox """" + Error(Err) + """" ' + Chr(13) + Chr(10) + "Try using the Data, Modify Values menu items to rescale variables with large values so that their magnitudes are decreased.", , "Error"
   Resume cpes
End If
cpes:

End Sub

Public Sub FiveP(cp1, cp2, ycp, ls, rs, cvrmse, rmse, r2)
ReDim A(numrecs, 3), beta(3, 1), sec(3), sigma(3)
Dim numints As Integer
Dim i As Integer, j As Integer, k As Integer
Dim numpasses As Integer
Dim xmid, gridpass, numleft, numright, indvar1, indvar2, numits
Dim rmsebest, r2best, cvrmsebest, cp1best, cp2best ', ycp, ls, rs
Dim totregs, numregs
On Error GoTo p5ce

'show perdone
Screen.MousePointer = 11
PerDone.Caption = "Processing Data..."
PerDone.Show 0
DoEvents
 
'calls SetAll to fill x(), y(), mins maxs etc.
'Call SetAll
'ReDim IndexNumO(n) As Integer, WDWEO(n) As Integer, XO(n), YO(n)

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo p5ce:
End If

modtype = "5P"

numints = 9
ReDim cpl(numints), cpr(numints)

'find total regressions to be performed for perdone
totregs = 16   'for fine search
For i = 1 To numints - 2   'for course search
   totregs = totregs + i
Next i

'sets pass
'If G1G2But.Value = True And gmodnum = 2 Then pass = 2 Else pass = 1

'refills x() and y() if appropriate
'If G1G2But.Value = True Then Call SetGrp

'temporary code to debug models
If 1 = 2 Then
ETMain.Graphbox.Cls
'ETMain.Graphbox.location
ETMain.Graphbox.CurrentX = -15
ETMain.Graphbox.CurrentY = 110
Open "c:\emodel\ettest.dat" For Output As #6
For i = 1 To n
   'ETMain.Graphbox.Print i, Format(x(i, 2), "0.00"), Format(y(i, 1), "0.00")
   Print #6, Format(x(i, 2), "0.00"), Format(y(i, 1), "0.00")
Next i
Close #6
'Stop
'GoTo p5ce
End If

'finds best-fit 5P model
For gridpass = 1 To 2

   'finds cps
   If gridpass = 1 Then 'course (first) grid search
      
      'calcs initial search grid
      inc = (xmax - xmin) / numints
      numleft = numints - 1
      numright = numints - 1
      For i = 1 To numints - 1
         cpl(i) = xmin + i * inc
         cpr(i) = xmin + i * inc
      Next i
      
   Else  'fine (fine) grid search
     
      'searches on either side of cp1 and cp2
      numleft = 4
      numright = 4
      cpl(1) = cp1best - 0.666 * inc
      cpl(2) = cp1best - 0.333 * inc
      cpl(3) = cp1best + 0.333 * inc
      cpl(4) = cp1best + 0.333 * inc
      cpr(1) = cp2best - 0.666 * inc
      cpr(2) = cp2best - 0.333 * inc
      cpr(3) = cp2best + 0.333 * inc
      cpr(4) = cp2best + 0.333 * inc
   End If
   
   'trys all combinations of change points
   For i = 1 To numleft
      For j = 1 To numright
                  
         'sets change points
         cp1 = cpl(i)
         cp2 = cpr(j)
         'performs regression
         If cp1 < cp2 Then
            numits = numits + 1
            For k = 1 To n
               If x(k, 2) <= cp1 Then
                  indvar1 = 1
               Else
                  indvar1 = 0
               End If
                  
               If x(k, 2) <= cp2 Then
                  indvar2 = 0
               Else
                  indvar2 = 1
               End If
               A(k, 1) = 1
               A(k, 2) = indvar1 * (x(k, 2) - cp1)
               A(k, 3) = indvar2 * (x(k, 2) - cp2)
            Next k
            'Call Reg(A(), n, 3, y(), n, 1, beta(), betar, betac, sec())
            'Call Inf(A(), n, 3, y(), n, 1, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), P)
            Call Reg(A(), n, 3, y(), n, 1, beta(), betar, betac, sec())
            Call Inf(A(), n, 3, y(), n, 1, beta(), betar, betac, sec(), rmse, cvrmse, r2, adjr2, sigma(), P)

            'Selects cps which give min RMSE
            If (gridpass = 1 And numits = 1) Or rmse < rmsebest Then
               rmsebest = rmse
               r2best = r2
               cvrmsebest = cvrmse
               cp1best = cp1
               cp2best = cp2
               ycp = beta(1, 1): ycpse = sigma(1)
               ls = beta(2, 1): lsse = sigma(2)
               rs = beta(3, 1): rsse = sigma(3)
            End If
         
            'update perdone
            regnum = regnum + 1
            Call UpdatePerDone(regnum, totregs, 1)

         End If
      Next j
   Next i
Next gridpass

'recalls reg and inf with final cps to fix residuals
'sets change points
cp1 = cp1best
cp2 = cp2best
'fills A() in accordance with cp1 and cp2
For k = 1 To n
   If x(k, 2) <= cp1 Then
      indvar1 = 1
   Else
      indvar1 = 0
   End If
      
   If x(k, 2) <= cp2 Then
      indvar2 = 0
   Else
      indvar2 = 1
   End If
   A(k, 1) = 1
   A(k, 2) = indvar1 * (x(k, 2) - cp1)
   A(k, 3) = indvar2 * (x(k, 2) - cp2)
Next k
'performs regression
'Call Reg(A(), n, 3, y(), n, 1, beta(), betar, betac, sec())
'Call Inf(A(), n, 3, y(), n, 1, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), P)
Call Reg(A(), n, 3, y(), n, 1, beta(), betar, betac, sec())
Call Inf(A(), n, 3, y(), n, 1, beta(), betar, betac, sec(), rmse, cvrmse, r2, adjr2, sigma(), P)

p5ce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume p5es
End If
p5es:

End Sub

Public Sub CalcSSNAC(xfld, yfld, yfldproj, grpfld)
Dim i, j
Dim ac, nac, startdate, enddate, numdayspost, numpostobs
Dim xcp, ycp, ls, rs, cp1, cp2
ReDim ivcoefs(8)
ReDim toanorm(numrecs), tiv1norm(numrecs), tiv2norm(numrecs)
Dim numdays, ttot, ttot2, wdate, numxvars

On Error GoTo csce
Screen.MousePointer = 11

'check if weather, energy and tm2 files are loaded
If numrecs = Empty Then
   msg$ = "Must open both energy and weather files before calculating NAC."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf eopen <> True Then
   msg$ = "Must open energy file before calculating NAC."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf wopen <> True Then
   msg$ = "Must open weather file before calculating NAC."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf tm2open <> True Then
   msg$ = "Must open TMY2 weather data file before calculating NAC."
   MsgBox msg$, , "Error"
   GoTo csce:
End If

'initialize counters
numdayspost = 0
numpostobs = 0
nac = 0
ac = 0

'numindvars determined in "open energy cmd" after reading utl file
numxvars = numindvars

If nmt = "2P" Then
   Call CalcSavings2PMVR(xfld, yfld, yfldproj, grpfld)
ElseIf nmt = "3PC" Or nmt = "3PH" Then
   If engytype$ = "elec" Then
      Index = 0 'for 3PC
   ElseIf engytype$ = "fuel" Then
      Index = 1 'for 3PH
   Else
      Stop
   End If
   Call CalcSavings3PMVR(Index, xfld, yfld, yfldproj, grpfld)
ElseIf nmt = "4P" Then
   Call CalcSavings4PMVR(xfld, yfld, yfldproj, grpfld)
ElseIf nmt = "5P" Then
   Call CalcSavings5PMVR(xfld, yfld, yfldproj, grpfld)
ElseIf nmt = "AS" Then
   Stop
End If

'fill Toanorm() from tmy2 data
startdate = CDate(Format(e(0, 1)) + "/" + Format(e(0, 2)) + "/" + Format(e(0, 3)))
For i = 1 To numrecs
   enddate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
   
   'calc normal temp during energy period from tmy2 avg daily temps
   If Year(startdate) = Year(enddate) Then
      numdays = 0
      ttot = 0
      For j = 1 To 365
         wdate = CDate(Format(w(j, 1)) + "/" + Format(w(j, 2)) + "/" + Format(Year(startdate)))
         If wdate > startdate And wdate <= enddate Then
            numdays = numdays + 1
            ttot = ttot + w(j, 4)
         End If
      Next j
      toanorm(i) = ttot / numdays
   Else 'startyear < endyear
      numdays = 0
      ttot = 0
      'calc sum of temps from startdate to end of calendar year
      For j = 1 To 365
         wdate = CDate(Format(w(j, 1)) + "/" + Format(w(j, 2)) + "/" + Format(Year(startdate)))
         If wdate > startdate Then
            numdays = numdays + 1
            ttot = ttot + w(j, 4)
         End If
      Next j
      'add sum of temps from start of calendar year to enddate
      For j = 1 To 365
         wdate = CDate(Format(w(j, 1)) + "/" + Format(w(j, 2)) + "/" + Format(Year(enddate)))
         If wdate < enddate Then
            numdays = numdays + 1
            ttot = ttot + w(j, 4)
         End If
      Next j
      'calc average temp during period
      toanorm(i) = ttot / numdays
   End If
   startdate = enddate
Next i
'test in immediate window   for i = 1 to numrecs: print i, toanorm(i), d(i,11): next i

'if numxvars > 1 then fill tiv1norm() and tiv2norm() from tiv() where tiv(i,1) = mo, tiv(i,2) = dy, tiv(i,3) = typindvar1, tiv(i,4) = typindvar2
If numxvars > 1 Then
   
   'check if tivfile is open
   If tivopen <> True Then
      msg$ = "Energy file (*.utl) includes independent variables, but typical independent variable file (*.tiv) is not open.  If single site analysis, Open TIV file after opening TM2 file.  If multisite analysis, add TIV file to Multisite List File."
      MsgBox msg$, , "Error"
      GoTo csce
   End If
 
   startdate = CDate(Format(e(0, 1)) + "/" + Format(e(0, 2)) + "/" + Format(e(0, 3)))
   For i = 1 To numrecs
      enddate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
      
      'calc normal temp during energy period from tmy2 avg daily temps
      If Year(startdate) = Year(enddate) Then
         numdays = 0
         ttot = 0
         ttot2 = 0
         For j = 1 To 365
            wdate = CDate(Format(tiv(j, 1)) + "/" + Format(tiv(j, 2)) + "/" + Format(Year(startdate)))
            If wdate > startdate And wdate <= enddate Then
               numdays = numdays + 1
               ttot = ttot + tiv(j, 3)
               If numxvars = 3 Then ttot2 = ttot2 + tiv(j, 4)
            End If
         Next j
         tiv1norm(i) = ttot / numdays
         If numxvars = 3 Then tiv2norm(i) = ttot2 / numdays
         
      Else 'startyear < endyear
         numdays = 0
         ttot = 0
         ttot2 = 0
         'calc sum of temps from startdate to end of calendar year
         For j = 1 To 365
            wdate = CDate(Format(tiv(j, 1)) + "/" + Format(tiv(j, 2)) + "/" + Format(Year(startdate)))
            If wdate > startdate Then
               numdays = numdays + 1
               ttot = ttot + tiv(j, 3)
               If numxvars = 3 Then ttot2 = ttot2 + tiv(j, 4)
            End If
         Next j
         'add sum of temps from start of calendar year to enddate
         For j = 1 To 365
            wdate = CDate(Format(tiv(j, 1)) + "/" + Format(tiv(j, 2)) + "/" + Format(Year(enddate)))
            If wdate < enddate Then
               numdays = numdays + 1
               ttot = ttot + tiv(j, 3)
               If numxvars = 3 Then ttot2 = ttot2 + tiv(j, 4)
            End If
         Next j
         'calc average temp during period
         tiv1norm(i) = ttot / numdays
         If numxvars = 3 Then tiv2norm(i) = ttot2 / numdays
      End If
      startdate = enddate
   Next i
End If
'test in immediate window For i = 1 To numrecs: Print i, tiv1norm(i), d(i, 9), tiv2norm(i), d(i, 10): Next i
      
startdate = CDate(Format(e(0, 1)) + "/" + Format(e(0, 2)) + "/" + Format(e(0, 3)))
For i = 1 To numrecs
   enddate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
   
   If d(i, xfld) <> ndflag And d(i, yfld) <> ndflag And d(i, grpfld) = 1 Then
      If numxvars = 1 Or (numxvars = 2 And d(i, 9) <> -99) Or (numxvars = 3 And d(i, 9) <> -99 And d(i, 10) <> -99) Then
         
         If nmt = "2P" Then
        
            'transfer global coefs to local coefs
            'xcp = nxcp1
            ycp = nycp
            ls = nls
            'rs = nrs
            ivcoefs(2) = nivcoefs(2)
            ivcoefs(3) = nivcoefs(3)
   
            'calc predicted energy
            If numxvars = 1 Then
               d(i, yfldproj) = ycp + ls * toanorm(i) ' + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
            ElseIf numxvars = 2 Then
               'd(i, yfldproj) = ycp + ls * toanorm(i) + ivcoefs(2) * d(i, 9)  '+ ivcoefs(3) * d(i, 10)
               d(i, yfldproj) = ycp + ls * toanorm(i) + ivcoefs(2) * tiv1norm(i)  '+ ivcoefs(3) * d(i, 10)
            ElseIf numxvars = 3 Then
               'd(i, yfldproj) = ycp + ls * toanorm(i) + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
               d(i, yfldproj) = ycp + ls * toanorm(i) + ivcoefs(2) * tiv1norm(i) + ivcoefs(3) * tiv2norm(i)
            End If
         
         ElseIf nmt = "3PC" Or nmt = "3PH" Then
            
            'transfer global coefs to local coefs
            xcp = nxcp1
            ycp = nycp
            ls = nls
            rs = nrs
            ivcoefs(2) = nivcoefs(2)
            ivcoefs(3) = nivcoefs(3)
   
            'calc predicted energy
            If toanorm(i) <= xcp Then
              indvar1 = 1
              indvar2 = 0
            Else
               indvar1 = 0
               indvar2 = 1
            End If
            If numxvars = 1 Then
               d(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) ' + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
            ElseIf numxvars = 2 Then
               'd(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * d(i, 9)  '+ ivcoefs(3) * d(i, 10)
               d(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * tiv1norm(i)  '+ ivcoefs(3) * d(i, 10)
            ElseIf numxvars = 3 Then
               'd(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
               d(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * tiv1norm(i) + ivcoefs(3) * tiv2norm(i)
            End If
         
         
         ElseIf nmt = "4P" Then
        
            'transfer global coefs to local coefs
            xcp = nxcp1
            ycp = nycp
            ls = nls
            rs = nrs
            ivcoefs(2) = nivcoefs(2)
            ivcoefs(3) = nivcoefs(3)
   
            'calc predicted energy
            If toanorm(i) <= xcp Then
              indvar1 = 1
              indvar2 = 0
            Else
               indvar1 = 0
               indvar2 = 1
            End If
            If numxvars = 1 Then
               d(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) ' + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
            ElseIf numxvars = 2 Then
               'd(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * d(i, 9)  '+ ivcoefs(3) * d(i, 10)
               d(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * tiv1norm(i)  '+ ivcoefs(3) * d(i, 10)
            ElseIf numxvars = 3 Then
               'd(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
               d(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * tiv1norm(i) + ivcoefs(3) * tiv2norm(i)
            End If
         
         ElseIf nmt = "5P" Then
            
            'transfer global coefs to local coefs
            cp1 = nxcp1
            cp2 = nxcp2
            ycp = nycp
            ls = nls
            rs = nrs
            ivcoefs(2) = nivcoefs(2)
            ivcoefs(3) = nivcoefs(3)
            
            'calc predicted energy
            If toanorm(i) <= cp1 Then
              indvar1 = 1
            Else
               indvar1 = 0
            End If
            If toanorm(i) <= cp2 Then
               indvar2 = 0
            Else
               indvar2 = 1
            End If
                     
            If numxvars = 1 Then
               d(i, yfldproj) = ycp + ls * indvar1 * (toanorm(i) - cp1) + rs * indvar2 * (toanorm(i) - cp2) ' + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
            ElseIf numxvars = 2 Then
               'd(i, yfldproj) = ycp + ls * indvar1 * (toanorm(i) - cp1) + rs * indvar2 * (toanorm(i) - cp2) + ivcoefs(2) * d(i, 9)  '+ ivcoefs(3) * d(i, 10)
               d(i, yfldproj) = ycp + ls * indvar1 * (toanorm(i) - cp1) + rs * indvar2 * (toanorm(i) - cp2) + ivcoefs(2) * tiv1norm(i)  '+ ivcoefs(3) * d(i, 10)
            ElseIf numxvars = 3 Then
               'd(i, yfldproj) = ycp + ls * indvar1 * (toanorm(i) - cp1) + rs * indvar2 * (toanorm(i) - cp2) + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
               d(i, yfldproj) = ycp + ls * indvar1 * (toanorm(i) - cp1) + rs * indvar2 * (toanorm(i) - cp2) + ivcoefs(2) * tiv1norm(i) + ivcoefs(3) * tiv2norm(i)
            End If
         
         
         ElseIf nmt = "AS" Then
            
            Stop
         
         End If
         
         'sum results
         numdayspost = numdayspost + (enddate - startdate)
         numpostobs = numpostobs + 1
         ac = ac + d(i, yfld)
         nac = nac + d(i, yfldproj)
         
      Else
         d(i, yfldproj) = ndflag
      End If
   Else
      d(i, yfldproj) = ndflag
   End If
   startdate = enddate
   
Next i

'adjusts totsav for other than 365.25 days
If numdayspost > 0 Then
   nac = nac * 365.25 / numdayspost
   ac = ac * 365.25 / numdayspost
Else
   nac = 0
   ac = 0
End If

'calc uncertainty
If nn > 0 And nac > 0 Then
   uncert = (1.96 * nrmse * ((1 + 2 / nn) * numpostobs) ^ 0.5) * 365.25 / numdayspost
   reluncert = Abs(uncert / nac) * 100
Else
   uncert = 0
   reluncert = 0
End If

'transfer results to global variables
nacg = nac
acg = ac
uncertg = uncert
reluncertg = reluncert

'graph projection
ETMain.Graphbox.Cls
Dim grid, datapnts
color1 = BLACK
color2 = BLUE
grid = True
datapnts = True

'modify variable name for display of NAC
If yfldproj = 12 Then
   vn$(12) = "Normal Elec Use (kWh/mo)"
ElseIf yfldproj = 13 Then
   vn$(13) = "Normal Elec Demand (kW)"
ElseIf yfldproj = 14 Then
   vn$(14) = "Normal Fuel Use (units/mo)"
End If
'call timeseries 2 graph
Call TS2Graph(yfld, yfldproj, grid, datapnts)
'reinstall original variable names
vn$(12) = "Elec Use if no retrofit (kWh/mo)"
vn$(13) = "Elec Demand if no retrofit (kW)"
vn$(14) = "Fuel Use if no retrofit (units/mo)"

'print results
'ETMain.StatusBox.Cls
'ETMain.StatusBox.Print "Pre-retrofit model: "; nmt; "   N = "; nn; "   R2 = "; Format(nr2, "0.00"); "   CV-RMSE = "; Format(ncvrmse, "0.0"); "%"
ETMain.StatusBox.Print "Annual consumption (AC) during baseline period, with actual weather = "; Format(ac, "#,##0.0"); " units/year"
ETMain.StatusBox.Print "Normal Annual Consumption (NAC) during baseline period, if period had normal (TMY2) weather = "; Format(nac, "#,##0.0"); " +- "; Format(uncert, "#,##0.0"); " ("; Format(reluncert, "#,##0.0"); "%)"; " units/year     % Change [(NAC-AC)/AC] = "; Format((nac - ac) / ac * 100, "#,##0.0"); "%"
'ETMain.StatusBox.Print " Number observations in NAC period = "; Format(numpostobs, "#,##0"); "  Number days in NAC period = "; Format(numdayspost, "#,##0")

Screen.MousePointer = 0
csce:
'display end message
Unload PerDone
'Close
Screen.MousePointer = 0
'prints error message
If Err Then
   msg$ = """" + Error(Err) + """"
   MsgBox msg$, , "Error"
   Resume cses
End If
cses:
End Sub

Public Sub ThreePMVR(x(), y(), n, numxvars, Index, xcp, sexcp, ycp, seycp, slope, seslope, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)
'Public Sub ThreePMVR(x(), y(), n, numxvars, Index, xcp, sexcp, ycp, seycp, slope, seslope, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)
Dim i, j, k, xmin, xmax, numints, inc, itration, IndVar, n1, n2
Dim betar, betac
Dim regnum, totregs 'for perdone
'sets global variable modtype

'turns on VB error handling
On Error GoTo psmce

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo psmce:
End If

'shows perdone
Screen.MousePointer = 11
PerDone.Caption = "Processing Data..."
PerDone.Show 0
DoEvents

'If index = 0 Then 3PCooling else 3PHeating
'modtype = "3P"

'dimensions arrays
'numxvars = 1
ReDim A(n, numxvars + 1), beta(numxvars + 1, 1), sec(numxvars + 1), sigma(numxvars + 1)

'finds x maxs and mins
For i = 1 To n
   If i = 1 Then
      xmin = x(1, 2)
      xmax = x(1, 2)
   End If
   If x(i, 2) < xmin Then xmin = x(i, 2)
   If x(i, 2) > xmax Then xmax = x(i, 2)
Next i

'calcs initial size of search grid
numints = 10
inc = (xmax - xmin) / numints

'finds total regressions to be performed for perdone
totregs = 2 * (numints - 1) + 1 'for fine search

'finds cp using course then fine grid searches
For itration = 1 To 2
   If itration > 1 Then
      xmin = bestcp - inc
      inc = 2 * inc / numints
   End If

   'tries various changepoints
   For i = 1 To numints - 1
      
      'inits counters
      n1 = 0: n2 = 0

      'fills A()
      cp = xmin + i * inc
      For j = 1 To n
         If x(j, 2) <= cp Then
            n1 = n1 + 1
            If Index = 0 Then '3pc
               IndVar = 0
            Else '3ph
               IndVar = 1
            End If
         Else
            n2 = n2 + 1
            If Index = 0 Then '3pc
               IndVar = 1
            Else '3ph
               IndVar = 0
            End If
         End If
         A(j, 1) = 1
         A(j, 2) = IndVar * (x(j, 2) - cp)
         
         'adds additional indep variables to A()
         For k = 1 To numxvars - 1
            A(j, 2 + k) = x(j, 2 + k)
         Next k

      Next j

      'calls regression engine
      Call Reg(A(), n, numxvars + 1, y(), n, 1, beta(), betar, betac, sec())
      Call Inf(A(), n, numxvars + 1, y(), n, 1, beta(), betar, betac, sec(), rmse, cvrmse, r2, adjr2, sigma(), roe)
     
      'records cp with minimum rmse
      If i = 1 Then
         rmsemin = rmse
         bestcp = cp
      ElseIf rmse < rmsemin Then
         rmsemin = rmse
         bestcp = cp
      End If
    
      'updates perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, 1)
   
   Next i
Next itration

'uses best cp to refill A() then rerun regression
n1 = 0: n2 = 0
cp = bestcp
For j = 1 To n
   If x(j, 2) <= cp Then
      n1 = n1 + 1
      If Index = 0 Then
         IndVar = 0
      Else
         IndVar = 1
      End If
   Else
      n2 = n2 + 1
      If Index = 0 Then
         IndVar = 1
      Else
         IndVar = 0
      End If
   End If
   A(j, 1) = 1
   A(j, 2) = IndVar * (x(j, 2) - cp)
   
   'adds additional indep variables to A()
   For k = 1 To numxvars - 1
      A(j, 2 + k) = x(j, 2 + k)
   Next k

Next j
Call Reg(A(), n, numxvars + 1, y(), n, 1, beta(), betar, betac, sec())
Call Inf(A(), n, numxvars + 1, y(), n, 1, beta(), betar, betac, sec(), rmse, cvrmse, r2, adjr2, sigma(), roe)

'sets changepoints and slope
'if index = 0 then 3PCooling else 3PHeating
xcp = bestcp
sexcp = inc
ycp = beta(1, 1)
seycp = sigma(1)
slope = beta(2, 1)
seslope = sigma(2)
'transfers ind var coefs to ivcoefs() and std error of ind var coefs to seivcoefs()
'if one ind var, then the iv coef is placed in ivcoefs(2)....
For k = 2 To numxvars
   ivcoefs(k) = beta(k + 1, 1)
   seivcoefs(k) = sigma(k + 1)
Next k

'calcs prediction interval
't = 1.96 + 2.7 / (n - 3)
'predint = t * rmse * (1 + 1 / n) ^ 0.5

'vb error handling
psmce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
  MsgBox """" + Error(Err) + """", , "Error"
  Resume psmes
End If
psmes:
End Sub

Public Sub FourPMVR(x(), y(), n, numxvars, xcp, sexcp, ycp, seycp, ls, sels, rs, sers, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)
Dim i, j, k, xmin, xmax, inc, numints
Dim itration, xmin4p, n1, n2, cp
Dim IndVar, rmsemin, bestcp, t, predint
Dim betar, betac
Dim totregs, regnum 'for perdone

'turns on VB error handling
On Error GoTo cpce

'show perdone
Screen.MousePointer = 11
PerDone.Caption = "Processing Data..."
PerDone.Show 0
DoEvents

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo cpce:
End If

'set modtype (global variable)
modtype = "4PMVR"

'dimension arrays
'numxvars = 1
ReDim A(n, numxvars + 2), beta(numxvars + 2, 1), sec(numxvars + 2), sigma(numxvars + 2)

'finds x maxs and mins
For i = 1 To n
   If i = 1 Then
      xmin = x(1, 2)
      xmax = x(1, 2)
   End If
   If x(i, 2) < xmin Then xmin = x(i, 2)
   If x(i, 2) > xmax Then xmax = x(i, 2)
Next i

'calcs initial size of search grid
numints = 10
inc = (xmax - xmin) / numints

'find total regressions to be performed for perdone
totregs = 2 * (numints - 1) + 1 'for fine search

'finds cp by using course then fine grid searches
For itration = 1 To 2
   
   'set xmin4p and inc depending on course or fine grid search
   If itration = 1 Then
      xmin4p = xmin
   Else '(itration = 2 for fine grid...
      xmin4p = bestcp - inc
      inc = 2 * inc / numints
   End If
   
   'tries various changepoints
   For i = 1 To numints - 1
      
      'inits counters
      n1 = 0: n2 = 0

      'fills A()
      cp = xmin4p + i * inc
      For j = 1 To n
         If x(j, 2) <= cp Then
            n1 = n1 + 1
            IndVar = 0
         Else
            n2 = n2 + 1
            IndVar = 1
         End If
         A(j, 1) = 1
         A(j, 2) = x(j, 2)
         A(j, 3) = IndVar * (x(j, 2) - cp)
         
         'adds additional indep variables to A()
         For k = 1 To numxvars - 1
            A(j, 3 + k) = x(j, 2 + k)
         Next k
      Next j

      'calls regression engine
      Call Reg(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec())
      Call Inf(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec(), rmse, cvrmse, r2, adjr2, sigma(), roe)
      
      'records cp with minimum rmse
      If i = 1 Then
         rmsemin = rmse
         bestcp = cp
      ElseIf rmse < rmsemin Then
         rmsemin = rmse
         bestcp = cp
      End If
      
      'updates perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, 1)
   
   Next i
Next itration

'uses best cp to fill A, then reruns regression with best cp
n1 = 0: n2 = 0
cp = bestcp
For j = 1 To n
   If x(j, 2) <= cp Then
      n1 = n1 + 1
      IndVar = 0
   Else
      n2 = n2 + 1
      IndVar = 1
   End If
   A(j, 1) = 1
   A(j, 2) = x(j, 2)
   A(j, 3) = IndVar * (x(j, 2) - cp)
   
   'adds additional indep variables to A()
   For k = 1 To numxvars - 1
      A(j, 3 + k) = x(j, 2 + k)
   Next k
Next j
Call Reg(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec())
Call Inf(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec(), rmse, cvrmse, r2, adjr2, sigma(), roe)

'calculates changepoints and slopes from beta() for yhat = ycp + ls(t-xcp)-  +  rs(t-xcp)+
xcp = bestcp
sexcp = inc
ycp = beta(1, 1) + beta(2, 1) * bestcp
If Abs(beta(2, 1)) > 0 And Abs(xcp) > 0 Then
   seycp = (sigma(1) ^ 2 + (beta(2, 1) * xcp) ^ 2 * ((sigma(2) / beta(2, 1)) ^ 2 + (sexcp / xcp) ^ 2)) ^ 0.5
Else
   seycp = sexcp
End If
ls = beta(2, 1)
sels = sigma(2)
rs = beta(2, 1) + beta(3, 1)
sers = (sigma(2) ^ 2 + sigma(3) ^ 2) ^ 0.5

'transfers ind var coefs to ivcoefs() and std error of ind var coefs to seivcoefs()
'if one ind var, then the iv coef is placed in ivcoefs(2)....
For k = 2 To numxvars
   ivcoefs(k) = beta(k + 2, 1)
   seivcoefs(k) = sigma(k + 2)
Next k

'calcs prediction interval
't = 1.96 + 2.7 / (n - 3)
'predint = t * rmse * (1 + 1 / n) ^ 0.5

'vb error handling
cpce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   MsgBox """" + Error(Err) + """"
   Resume cpes
End If
cpes:

End Sub

Public Sub FivePMVR(x(), y(), n, numxvars, xcp1, sexcp1, xcp2, sexcp2, ycp, seycp, ls, sels, rs, sers, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)
'Public Sub FivePMVR(x(), y(), n, numxvars, xcp1, sexcp1, xcp2, sexcp2, ycp, seycp, ls, sels, rs, sers, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)
'change this transpose ls and rs in argument list
Dim i, j, k, m, xmin, xmax, numints
Dim gridpass, numleft, numright, indvar1, indvar2, rmsebest, cp1best, cp2best
Dim betar, betac
Dim totregs, regnum 'for perdone

'turns on VB error handling
On Error GoTo p5ce

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo p5ce:
End If

'show perdone
Screen.MousePointer = 11
PerDone.Caption = "Processing Data..."
PerDone.Show 0
DoEvents
 
'sets modtype
modtype = "5PMVR"

'dims arrays
ReDim A(n, numxvars + 2), beta(numxvars + 2, 1), sec(numxvars + 2), sigma(numxvars + 2)

'sets number of search intervals
numints = 9
ReDim cpl(numints), cpr(numints)

'find total regressions to be performed for perdone
totregs = 16   'for fine search
For i = 1 To numints - 2   'for course search
   totregs = totregs + i
Next i

'finds x maxs and mins
For i = 1 To n
   If i = 1 Then
      xmin = x(1, 2)
      xmax = x(1, 2)
   End If
   If x(i, 2) < xmin Then xmin = x(i, 2)
   If x(i, 2) > xmax Then xmax = x(i, 2)
Next i

'finds best changepoints using course then fine grid search
For gridpass = 1 To 2

   'finds change points
   If gridpass = 1 Then 'course (first) grid search
      
      'calcs initial search grid
      inc = (xmax - xmin) / numints
      numleft = numints - 1
      numright = numints - 1
      For i = 1 To numints - 1
         cpl(i) = xmin + i * inc
         cpr(i) = xmin + i * inc
      Next i
      
   Else  'fine (fine) grid search
     
      'searches on either side of xcp1 and xcp2
      numleft = 4
      numright = 4
      cpl(1) = cp1best - 0.666 * inc
      cpl(2) = cp1best - 0.333 * inc
      cpl(3) = cp1best + 0.333 * inc
      cpl(4) = cp1best + 0.666 * inc
      cpr(1) = cp2best - 0.666 * inc
      cpr(2) = cp2best - 0.333 * inc
      cpr(3) = cp2best + 0.333 * inc
      cpr(4) = cp2best + 0.666 * inc
   End If
   
   'trys all combinations of change points
   For i = 1 To numleft
      For j = 1 To numright
                  
         'sets change points
         xcp1 = cpl(i)
         xcp2 = cpr(j)
         
         'performs regression
         If xcp1 < xcp2 Then
            numits = numits + 1
            For k = 1 To n
               If x(k, 2) <= xcp1 Then
                  indvar1 = 1
               Else
                  indvar1 = 0
               End If
                  
               If x(k, 2) <= xcp2 Then
                  indvar2 = 0
               Else
                  indvar2 = 1
               End If
               A(k, 1) = 1
               A(k, 2) = indvar1 * (x(k, 2) - xcp1)
               A(k, 3) = indvar2 * (x(k, 2) - xcp2)
               
               'adds additional indep variables to A()
               For m = 1 To numxvars - 1
                  A(k, 3 + m) = x(k, 2 + m)
               Next m
            Next k
            Call Reg(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec())
            Call Inf(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec(), rmse, cvrmse, r2, adjr2, sigma(), P)

            'selects changepoints which give minimum RMSE
            If (gridpass = 1 And numits = 1) Or rmse < rmsebest Then
               rmsebest = rmse
               cp1best = xcp1
               cp2best = xcp2
            End If
            
            'update perdone
            regnum = regnum + 1
            Call UpdatePerDone(regnum, totregs, 1)

         End If
         
      Next j
   Next i
   
Next gridpass

'sets final change points equal to best-fit changepoints
xcp1 = cp1best
xcp2 = cp2best
sexcp1 = inc * 0.333
sexcp2 = inc * 0.333

'fills A() in accordance with xcp1 and xcp2
For k = 1 To n
   If x(k, 2) <= xcp1 Then
      indvar1 = 1
   Else
      indvar1 = 0
   End If
      
   If x(k, 2) <= xcp2 Then
      indvar2 = 0
   Else
      indvar2 = 1
   End If
   A(k, 1) = 1
   A(k, 2) = indvar1 * (x(k, 2) - xcp1)
   A(k, 3) = indvar2 * (x(k, 2) - xcp2)
   
   'adds additional indep variables to A()
   For m = 1 To numxvars - 1
      A(k, 3 + m) = x(k, 2 + m)
   Next m

Next k

'performs regression with best changepoints
Call Reg(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec())
Call Inf(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec(), rmse, cvrmse, r2, adjr2, sigma(), P)

'sets standard error of coefficents
ycp = beta(1, 1)
seycp = sigma(1)
ls = beta(2, 1)
sels = sigma(2)
rs = beta(3, 1)
sers = sigma(3)

'transfers ind var coefs to ivcoefs() and std error of ind var coefs to seivcoefs()
'if one ind var, then the iv coef is placed in ivcoefs(2)....
For k = 2 To numxvars
   ivcoefs(k) = beta(k + 2, 1)
   seivcoefs(k) = sigma(k + 2)
Next k

p5ce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume p5es
End If
p5es:

End Sub

Public Sub CalcSavingsASMVR(xfld, yfld, yfldproj, grpfld)
Dim i, j, numvars, numxvars
Dim rmse, cvrmse, r2, adjr2, roe, uncert, reluncert
Dim totsav, startdate, enddate, numdayspost, numpostobs
Dim xcp, cp1, cp2, ycp, ls, rs
Dim rmse4, rmse5, cvrmse4, cvrmse5, r24, r25
Dim modtype, etype, eval
Dim outfilename$
Dim totadjbase, totactpost
Dim percentsav, savperobs, savperday
Dim ytot, meanactpre, meanactpost
ReDim nivcoefs(8)

On Error GoTo csce
Screen.MousePointer = 11

'check if weather file is loaded
If numrecs = Empty Then
   msg$ = "Must open both energy and weather files before calculating savings."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf eopen <> True Then
   msg$ = "Must open energy file before calculating savings."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf wopen <> True Then
   msg$ = "Must open weather file before calculating savings."
   MsgBox msg$, , "Error"
   GoTo csce:
End If

'numindvars determined in "open energy cmd" after reading utl file
numxvars = numindvars

'dimension arrays and variables for calling cp-mvr models
ReDim x(numrecs, numxvars + 1), y(numrecs, numxvars + 1)
ReDim coef(numxvars + 1, 1), sec(numxvars + 1), sigma(numxvars + 1)
'ReDim ivcoefs(numxvars + 1), seivcoefs(numxvars + 1) 'note for cp-mvr models, coef for first ind var = ivcoefs(2), coef for second ind var = ivcoefs(3), etc.
ReDim ivcoefs(8), seivcoefs(8) 'note for cp-mvr models, coef for first ind var = ivcoefs(2), coef for second ind var = ivcoefs(3), etc.
Dim sexcp, seycp, sels, sers

'fill n, x(), y()
n = 0
meanactpre = 0
For i = 1 To numrecs
   If d(i, xfld) <> -99 And d(i, yfld) <> -99 And d(i, grpfld) = 1 Then
      If numxvars = 1 Or (numxvars = 2 And d(i, 9) <> -99) Or (numxvars = 3 And d(i, 9) <> -99 And d(i, 10) <> -99) Then
         'fill n, x() and y()
         n = n + 1
         x(n, 1) = 1
         x(n, 2) = d(i, xfld)
         If numxvars = 2 Then
            x(n, 3) = d(i, 9)
         ElseIf numxvars = 3 Then
            x(n, 3) = d(i, 9)
            x(n, 4) = d(i, 10)
         End If
         y(n, 1) = d(i, yfld)
         ytot = ytot + y(n, 1)
         meanactpre = ytot / n
   
         
         'finds max and mins
         If n = 1 Then
            xmin = x(1, 2)
            xmax = x(1, 2)
            ymin = y(1, 1)
            ymax = y(1, 1)
         End If
         If x(n, 2) < xmin Then xmin = x(n, 2)
         If x(n, 2) > xmax Then xmax = x(n, 2)
         If y(n, 1) < ymin Then ymin = y(n, 1)
         If y(n, 1) > ymax Then ymax = y(n, 1)
      End If
   End If
Next i


'select 4p or 5p model
Call FivePMVR(x(), y(), n, numxvars, cp1, sexcp1, cp2, sexcp2, ycp, seycp, ls, sels, rs, sers, ivcoefs(), seivcoefs(), rmse5, cvrmse5, r25)

If (ls < 0 And rs > 0) Or (ls > 0 And rs < 0) Then 'use 5p
   
   'set inf params
   modtype = "5P"
   cvrmse = cvrmse5
   rmse = rmse5
   r2 = r25
   
   'init counting variables
   numdayspost = 0
   numpostobs = 0
   totadjbase = 0
   totactpost = 0
   meanactpost = 0
   totsav = 0
   
   'fill d() with projected baseline model
   For i = 1 To numrecs
      If d(i, grpfld) = 1 Then 'pre
         d(i, yfldproj) = d(i, yfld)
         startdate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
      Else 'post
         enddate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
         If d(i, xfld) <> ndflag And d(i, yfld) <> ndflag Then
            If numxvars = 1 Or (numxvars = 2 And d(i, 9) <> -99) Or (numxvars = 3 And d(i, 9) <> -99 And d(i, 10) <> -99) Then
    
               'calc projected energy use
               If d(i, xfld) <= cp1 Then
                 indvar1 = 1
               Else
                  indvar1 = 0
               End If
               If d(i, xfld) <= cp2 Then
                  indvar2 = 0
               Else
                  indvar2 = 1
               End If
               If numxvars = 1 Then
                  d(i, yfldproj) = ycp + ls * indvar1 * (d(i, xfld) - cp1) + rs * indvar2 * (d(i, xfld) - cp2) ' + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
               ElseIf numxvars = 2 Then
                  d(i, yfldproj) = ycp + ls * indvar1 * (d(i, xfld) - cp1) + rs * indvar2 * (d(i, xfld) - cp2) + ivcoefs(2) * d(i, 9) '+ ivcoefs(3) * d(i, 10)
               ElseIf numxvars = 3 Then
                  d(i, yfldproj) = ycp + ls * indvar1 * (d(i, xfld) - cp1) + rs * indvar2 * (d(i, xfld) - cp2) + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
               End If
               
               'calc post days and savings
               numdayspost = numdayspost + (enddate - startdate)
               numpostobs = numpostobs + 1
               totadjbase = totadjbase + d(i, yfldproj)
               totactpost = totactpost + d(i, yfld)
               meanactpost = totactpost / numpostobs
               totsav = totsav + (d(i, yfldproj) - d(i, yfld))
            End If
         Else
            d(i, yfldproj) = ndflag
         End If
         startdate = enddate
      End If
   Next i

Else 'choose 4p model
   
   'call model
   Call FourPMVR(x(), y(), n, numxvars, xcp, sexcp, ycp, seycp, ls, sels, rs, sers, ivcoefs(), seivcoefs(), rmse4, cvrmse4, r24)
         
   'set inf params
   cvrmse = cvrmse4
   r2 = r24
   rmse = rmse4
   modtype = "4P"
   
   'init counting variables
   numdayspost = 0
   numpostobs = 0
   totadjbase = 0
   totactpost = 0
   meanactpost = 0
   totsav = 0
      
   'fill d( ) with projected baseline model
   For i = 1 To numrecs
      If d(i, grpfld) = 1 Then 'pre
         d(i, yfldproj) = d(i, yfld)
         startdate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
      Else 'post
         enddate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
         If d(i, xfld) <> ndflag And d(i, yfld) <> ndflag Then
            If numxvars = 1 Or (numxvars = 2 And d(i, 9) <> -99) Or (numxvars = 3 And d(i, 9) <> -99 And d(i, 10) <> -99) Then

                        
               'calc projected energy use
               If d(i, xfld) <= xcp Then
                 indvar1 = 1
                 indvar2 = 0
               Else
                  indvar1 = 0
                  indvar2 = 1
               End If
               If numxvars = 1 Then
                  d(i, yfldproj) = ycp - ls * indvar1 * (xcp - d(i, xfld)) + rs * indvar2 * (d(i, xfld) - xcp) ' + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
               ElseIf numxvars = 2 Then
                  d(i, yfldproj) = ycp - ls * indvar1 * (xcp - d(i, xfld)) + rs * indvar2 * (d(i, xfld) - xcp) + ivcoefs(2) * d(i, 9)  '+ ivcoefs(3) * d(i, 10)
               ElseIf numxvars = 3 Then
                  d(i, yfldproj) = ycp - ls * indvar1 * (xcp - d(i, xfld)) + rs * indvar2 * (d(i, xfld) - xcp) + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
               End If
               
               'calc post days and savings
               numdayspost = numdayspost + (enddate - startdate)
               numpostobs = numpostobs + 1
               totadjbase = totadjbase + d(i, yfldproj)
               totactpost = totactpost + d(i, yfld)
               totsav = totsav + (d(i, yfldproj) - d(i, yfld))
            End If
         Else
            d(i, yfldproj) = ndflag
         End If
         startdate = enddate
      End If
   Next i
End If

'calc uncertainty
If n > 0 And totsav <> 0 Then
   uncert = 1.96 * rmse * ((1 + 2 / n) * numpostobs) ^ 0.5
   reluncert = Abs(uncert / totsav) * 100
Else
   uncert = -99
   reluncert = -99
   totsav = 0
End If

'calc some savings metrics
If totactpost > 0 Then percentsav = totsav / totactpost * 100 Else percentsav = 0
If numpostobs > 0 Then savperobs = totsav / numpostobs Else savperobs = 0
If numdayspost > 0 Then savperday = totsav / numdayspost Else savperday = 0

'graph projection
ETMain.Graphbox.Cls
Dim grid, datapnts
color1 = BLACK
color2 = BLUE
grid = True
datapnts = True
Call TS2Graph(yfld, yfldproj, grid, datapnts)

'set energy type and output filename
If yfld = 4 Then
   etype = "Elec Use"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "kwh"
   eval = 1
ElseIf yfld = 5 Then
   etype = "Elec Demand"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "kwd"
   eval = 2
ElseIf yfld = 6 Then
   etype = "Fuel Use"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "thm"
   eval = 3
Else
   etype = "Energy"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "out"
End If

'print results
ETMain.StatusBox.Cls
Open outfilename$ For Output As #1
If modtype = "4P" Then
   'print results to status box
   If numxvars = 1 Then
      ETMain.StatusBox.Print "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00") '; "   IV1c = "; Format(ivcoefs(2), "#,##0.0000");"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
      ETMain.StatusBox.Print "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(xcp, "#,##0.00"); ")+"
   ElseIf numxvars = 2 Then
      ETMain.StatusBox.Print "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000"); '"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
      ETMain.StatusBox.Print "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(xcp, "#,##0.00"); ")+  + "; Format(ivcoefs(2), "#,##0.00"); " IV1"
   ElseIf numxvars = 3 Then
      ETMain.StatusBox.Print "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000"); "   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
      ETMain.StatusBox.Print "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(xcp, "#,##0.00"); ")+  + "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
   End If
               
   'print results to output file
   If numxvars = 1 Then
      Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00") '; "   IV1c = "; Format(ivcoefs(2), "#,##0.0000");"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
      Print #1, "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(xcp, "#,##0.00"); ")+"
   ElseIf numxvars = 2 Then
      Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000"); '"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
      Print #1, "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(xcp, "#,##0.00"); ")+  + "; Format(ivcoefs(2), "#,##0.00"); " IV1"
   ElseIf numxvars = 3 Then
      Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000"); "   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
      Print #1, "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(xcp, "#,##0.00"); ")+  + "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
   End If
   
   'initialize global reg parameter variables
   nxcp1 = -99
   nxcp2 = -99
   nycp = -99
   nls = -99
   nrs = -99
   
   'save reg params as global variables for nac calc
   nmt = modtype
   nxcp1 = xcp
   nxcp2 = -99
   nycp = ycp
   nls = ls
   nrs = rs
   nivcoefs(2) = ivcoefs(2)
   nivcoefs(3) = ivcoefs(3)
   
   'fill coef matrix for drawing model lines on xy plots
   For i = 1 To 3
      For j = 1 To 5
         c(i, j) = -99
      Next j
   Next i
   c(eval, 2) = xcp
   c(eval, 3) = ycp
   c(eval, 4) = ls
   c(eval, 5) = rs
   
Else '5p model
   
   'print results to status box
   If numxvars = 1 Then
      ETMain.StatusBox.Print "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp1 = "; Format(cp1, "#,##0.00"); "   Xcp2 = "; Format(cp2, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00") '; "   IV1c = "; Format(ivcoefs(2), "#,##0.0000");"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
      ETMain.StatusBox.Print "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(cp1, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(cp2, "#,##0.00"); ")+" ' + "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
   ElseIf numxvars = 2 Then
      ETMain.StatusBox.Print "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp1 = "; Format(cp1, "#,##0.00"); "   Xcp2 = "; Format(cp2, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000") ';"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
      ETMain.StatusBox.Print "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(cp1, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(cp2, "#,##0.00"); ")+  + "; Format(ivcoefs(2), "#,##0.00"); " IV1" ';  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
   ElseIf numxvars = 3 Then
      ETMain.StatusBox.Print "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp1 = "; Format(cp1, "#,##0.00"); "   Xcp2 = "; Format(cp2, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000"); "   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
      ETMain.StatusBox.Print "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(cp1, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(cp2, "#,##0.00"); ")+  + "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
   End If
   
   'print results to output file
   If numxvars = 1 Then
      Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp1 = "; Format(cp1, "#,##0.00"); "   Xcp2 = "; Format(cp2, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00") '; "   IV1c = "; Format(ivcoefs(2), "#,##0.0000");"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
      Print #1, "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(cp1, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(cp2, "#,##0.00"); ")+" ' + "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
   ElseIf numxvars = 2 Then
      Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp1 = "; Format(cp1, "#,##0.00"); "   Xcp2 = "; Format(cp2, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000") ';"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
      Print #1, "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(cp1, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(cp2, "#,##0.00"); ")+  + "; Format(ivcoefs(2), "#,##0.00"); " IV1" ';  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
   ElseIf numxvars = 3 Then
      Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp1 = "; Format(cp1, "#,##0.00"); "   Xcp2 = "; Format(cp2, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000"); "   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
      Print #1, "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(cp1, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(cp2, "#,##0.00"); ")+  + "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
   End If
   
   'initialize global reg parameter variables
   nxcp1 = -99
   nxcp2 = -99
   nycp = -99
   nls = -99
   nrs = -99

   'save reg params as global variables for nac calc
   nmt = modtype
   nxcp1 = cp1
   nxcp2 = cp2
   nycp = ycp
   nls = ls
   nrs = rs
   nivcoefs(2) = ivcoefs(2)
   nivcoefs(3) = ivcoefs(3)
   
   'fill coef matrix for drawing model lines on xy plots
   For i = 1 To 3
      For j = 1 To 5
         c(i, j) = -99
      Next j
   Next i
   c(eval, 1) = cp1
   c(eval, 2) = cp2
   c(eval, 3) = ycp
   c(eval, 4) = ls
   c(eval, 5) = rs
End If


'print savings summary to screen
If totactpost > 0 And numpostobs > 0 And numdayspost > 0 Then
   ETMain.StatusBox.Print "Adjusted Baseline Use During Post-retrofit Period = "; Tab(60); Format(totadjbase, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   ETMain.StatusBox.Print "Actual Use During Post-retrofit Period = "; Tab(60); Format(totactpost, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   ETMain.StatusBox.Print "Total Savings During Post-retrofit Period = "; Tab(60); Format(totsav, "#,##0"); " ("; Format(percentsav, "0"); "%) +- "; Format(uncert, "#,##0.0"); " ("; Format(reluncert, "#,##0.0"); "%) over "; numpostobs; " observations and "; numdayspost; " days"
   ETMain.StatusBox.Print "Average Savings During Post-retrofit Period = "; Tab(60); Format(savperobs, "#,##0.00"); " per observation   =   "; Format(savperday, "#,##0.000"); " per day"
End If

'print savings summary to file
If totactpost > 0 And numpostobs > 0 And numdayspost > 0 Then
   Print #1, "Adjusted Baseline Use During Post-retrofit Period = "; Tab(60); Format(totadjbase, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   Print #1, "Actual Use During Post-retrofit Period = "; Tab(60); Format(totactpost, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   Print #1, "Total Savings During Post-retrofit Period = "; Tab(60); Format(totsav, "#,##0"); " ("; Format(percentsav, "0"); "%) +- "; Format(uncert, "#,##0.0"); " ("; Format(reluncert, "#,##0.0"); "%) over "; numpostobs; " observations and "; numdayspost; " days"
   Print #1, "Average Savings During Post-retrofit Period = "; Tab(60); Format(savperobs, "#,##0.00"); " per observation   =   "; Format(savperday, "#,##0.000"); " per day"
End If
Close #1

'print output data *.dat
outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "dat"
Open outfilename$ For Output As #1
For i = 1 To numrecs
   For j = 1 To 8
      Print #1, d(i, j),
   Next j
   For j = 9 To 12
      Print #1, Format(d(i, j), "#,##0.000"),
   Next j
   Print #1,
Next i
Close #1

'save reg params as global variables for nac calculations
nrmse = rmse
nn = n
ncvrmse = cvrmse
nr2 = r2

'save savings metrics as global variables for multilist output
nnumdayspost = numdayspost
nnumpostobs = numpostobs
ntotsav = totsav
nuncert = uncert
npercentsav = percentsav
nreluncert = reluncert
nsavperobs = savperobs
nsavperday = savperday
nmeanactpre = meanactpre
nmeanactpost = meanactpost

csce:
'display end message
Unload PerDone
Close #1
Screen.MousePointer = 0
'prints error message
If Err Then
   msg$ = """" + Error(Err) + """"
   MsgBox msg$, , "Error"
   Resume cses
End If
cses:

End Sub

Public Sub FillDText(filename$)
Dim l$, delim, charnum
Dim i, j, scharnum, fnum, echarnum, flength, row, col

'sets default no data flag
ndflag = -99

'finds delimiter type, numflds, numrecs
numrecs = 0
numflds = 0
Open filename$ For Input As #1
While Not EOF(1)
   'reads a line
   Line Input #1, l$
   l$ = Trim$(l$)
   'If (Len(l$) > 0) And (Left(l$, 1) = "1" Or Left(l$, 1) = "2" Or Left(l$, 1) = "3" Or Left(l$, 1) = "4" Or Left(l$, 1) = "5" Or Left(l$, 1) = "6" Or Left(l$, 1) = "7" Or Left(l$, 1) = "8" Or Left(l$, 1) = "9" Or Left(l$, 1) = "0") Then
   If (Len(l$) > 0) Then
      numrecs = numrecs + 1
   
      If numrecs = 1 Then
         
         'determines if comma, tab (chr(9)), or space (chr(32)) delimited
         delim = ""
         For charnum = 1 To Len(l$)
            If Mid$(l$, charnum, 1) = "," Then
               delim = ","
               Exit For
            End If
         Next charnum
         If delim <> "," Then
            For charnum = 1 To Len(l$)
               If Mid$(l$, charnum, 1) = Chr(9) Then
                  delim = Chr(9)
                  Exit For
               End If
            Next charnum
         End If
         If delim = "" Then delim = Chr(32)
   
         'determines numflds and dims field()
         numflds = 0
         If delim = "," Or delim = Chr(9) Then
            For charnum = 1 To Len(l$)
               If Mid$(l$, charnum, 1) = delim Or charnum = Len(l$) Then
                  numflds = numflds + 1
               End If
            Next charnum
         Else 'delim = " "
            For charnum = 1 To Len(l$)
               If Mid$(l$, charnum, 1) = delim Then
                  If Mid$(l$, charnum - 1, 1) <> delim Then numflds = numflds + 1
               ElseIf charnum = Len(l$) Then
                  numflds = numflds + 1
               End If
            Next charnum
         End If
      End If
   End If
Wend
Close #1
ReDim d(numrecs, numflds)


'fills d(row,col) if space, tab or comma delim
row = 0
Open filename$ For Input As #1
While Not EOF(1)
   'reads a line
   Line Input #1, l$
   l$ = Trim$(l$)
   'If (Len(l$) > 0) And (Left(l$, 1) = "1" Or Left(l$, 1) = "2" Or Left(l$, 1) = "3" Or Left(l$, 1) = "4" Or Left(l$, 1) = "5" Or Left(l$, 1) = "6" Or Left(l$, 1) = "7" Or Left(l$, 1) = "8" Or Left(l$, 1) = "9" Or Left(l$, 1) = "0") Then
   If (Len(l$) > 0) Then
      row = row + 1
      'parses the line into fields
      scharnum = 1
      fnum = 0
      For charnum = 1 To Len(l$)
         If Mid$(l$, charnum, 1) = delim Then
            If Mid$(l$, charnum - 1, 1) <> delim Then
               fnum = fnum + 1
               If fnum > numflds Then
                  msg$ = "Incorrect number of columns in line " + Str$(row) + " of " + UCase$(filename$) + "."
                  MsgBox msg$, , "Error"
                  Close #1
                  Screen.MousePointer = 0
                  Exit Sub
               End If
               echarnum = charnum - 1
               flength = echarnum - scharnum + 1
               'd(row, fnum) = Val(Trim(Mid$(l$, scharnum, flength)))
               d(row, fnum) = Trim(Mid$(l$, scharnum, flength))
               scharnum = echarnum + 2
            End If
         ElseIf charnum = Len(l$) Then
            fnum = fnum + 1
            echarnum = charnum
            flength = echarnum - scharnum + 1
            'd(row, fnum) = Val(Trim(Mid$(l$, scharnum, flength)))
            d(row, fnum) = Trim(Mid$(l$, scharnum, flength))
         End If
      Next charnum
   End If
   Call UpdatePerDone(row, numrecs * 2, 1)
Wend

Close #1

End Sub


Public Sub CalcSavings3PMVR(Index, xfld, yfld, yfldproj, grpfld)
Dim i, j, numvars, numxvars
Dim rmse, cvrmse, r2, adjr2, roe, uncert, reluncert
Dim totsav, startdate, enddate, numdayspost, numpostobs
Dim xcp, cp1, cp2, ycp, ls, rs, ivar
'Dim rmse4, rmse5, cvrmse4, cvrmse5, r24, r25
Dim modtype, etype, eval
Dim outfilename$
Dim totadjbase, totactpost
Dim percentsav, savperobs, savperday
Dim ytot, meanactpre, meanactpost
ReDim nivcoefs(8)

On Error GoTo csce
Screen.MousePointer = 11

'check if weather file is loaded
If numrecs = Empty Then
   msg$ = "Must open both energy and weather files before calculating savings."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf eopen <> True Then
   msg$ = "Must open energy file before calculating savings."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf wopen <> True Then
   msg$ = "Must open weather file before calculating savings."
   MsgBox msg$, , "Error"
   GoTo csce:
End If

'numindvars determined in "open energy cmd" after reading utl file
numxvars = numindvars

'dimension arrays and variables for calling cp-mvr models
ReDim x(numrecs, numxvars + 1), y(numrecs, numxvars + 1)
ReDim coef(numxvars + 1, 1), sec(numxvars + 1), sigma(numxvars + 1)
'ReDim ivcoefs(numxvars + 1), seivcoefs(numxvars + 1) 'note for cp-mvr models, coef for first ind var = ivcoefs(2), coef for second ind var = ivcoefs(3), etc.
ReDim ivcoefs(8), seivcoefs(8) 'note for cp-mvr models, coef for first ind var = ivcoefs(2), coef for second ind var = ivcoefs(3), etc.
Dim sexcp, seycp, sels, sers

'fill n, x(), y()
n = 0
meanactpre = 0
For i = 1 To numrecs
   If d(i, xfld) <> -99 And d(i, yfld) <> -99 And d(i, grpfld) = 1 Then
      If numxvars = 1 Or (numxvars = 2 And d(i, 9) <> -99) Or (numxvars = 3 And d(i, 9) <> -99 And d(i, 10) <> -99) Then
      
         'fill n, x() and y()
         n = n + 1
         x(n, 1) = 1
         x(n, 2) = d(i, xfld)
         If numxvars = 2 Then
            x(n, 3) = d(i, 9)
         ElseIf numxvars = 3 Then
            x(n, 3) = d(i, 9)
            x(n, 4) = d(i, 10)
         End If
         y(n, 1) = d(i, yfld)
         ytot = ytot + y(n, 1)
         meanactpre = ytot / n
         
         'finds max and mins
         If n = 1 Then
            xmin = x(1, 2)
            xmax = x(1, 2)
            ymin = y(1, 1)
            ymax = y(1, 1)
         End If
         If x(n, 2) < xmin Then xmin = x(n, 2)
         If x(n, 2) > xmax Then xmax = x(n, 2)
         If y(n, 1) < ymin Then ymin = y(n, 1)
         If y(n, 1) > ymax Then ymax = y(n, 1)
      
      End If
   End If
Next i

'calls 3p-mvr model to determine coefs and statistics
rmse = -99
cvrmse = -99
r2 = -99
Call ThreePMVR(x(), y(), n, numxvars, Index, xcp, sexcp, ycp, seycp, slope, seslope, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)
'sets right and left slopes
If Index = 0 Then '3pc
   ls = 0: sels = 0
   rs = slope: sers = seslope
   modtype = "3PC"
Else '3ph
   ls = slope: sels = seslope
   rs = 0: sers = 0
   modtype = "3PH"
End If
modeleqn = Format(ycp, "#,##0.00") + "  -  " + Format(ls, "#,##0.00") + " (" + Format(xcp, "#,##0.00") + " - " + "T" + ")+  +  " + Format(rs, "#,##0.00") + " (" + "T" + " - " + Format(xcp, "#,##0.00") + ")+"

'init counting variables
numdayspost = 0
numpostobs = 0
totadjbase = 0
totactpost = 0
meanactpost = 0
totsav = 0
   
'fill d( ) with projected baseline model

For i = 1 To numrecs
   If d(i, grpfld) = 1 Then 'pre
      d(i, yfldproj) = d(i, yfld)
      startdate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
   Else 'post
      enddate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
      If d(i, xfld) <> ndflag And d(i, yfld) <> ndflag Then
         If numxvars = 1 Or (numxvars = 2 And d(i, 9) <> -99) Or (numxvars = 3 And d(i, 9) <> -99 And d(i, 10) <> -99) Then
         
            'calc projected energy use
            If Index = 0 Then '3pc
               If d(i, xfld) <= xcp Then ivar = 0 Else ivar = 1
            Else '3ph
               If d(i, xfld) <= xcp Then ivar = 1 Else ivar = 0
            End If
            If numxvars = 1 Then
               d(i, yfldproj) = ycp + slope * (d(i, xfld) - xcp) * ivar
            ElseIf numxvars = 2 Then
               d(i, yfldproj) = ycp + slope * (d(i, xfld) - xcp) * ivar + ivcoefs(2) * d(i, 9)
            ElseIf numxvars = 3 Then
               d(i, yfldproj) = ycp + slope * (d(i, xfld) - xcp) * ivar + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
            End If
            
            'calc post days and savings
            numdayspost = numdayspost + (enddate - startdate)
            numpostobs = numpostobs + 1
            totadjbase = totadjbase + d(i, yfldproj)
            totactpost = totactpost + d(i, yfld)
            meanactpost = totactpost / numpostobs
            totsav = totsav + (d(i, yfldproj) - d(i, yfld))
         
         End If
      Else
         d(i, yfldproj) = ndflag
      End If
      startdate = enddate
   End If
Next i


'calc uncertainty
If n > 0 And totsav <> 0 Then
   uncert = 1.96 * rmse * ((1 + 2 / n) * numpostobs) ^ 0.5
   reluncert = Abs(uncert / totsav) * 100
Else
   uncert = -99
   reluncert = -99
   totsav = 0
End If

'calc some savings metrics
If totactpost > 0 Then percentsav = totsav / totactpost * 100 Else percentsav = 0
If numpostobs > 0 Then savperobs = totsav / numpostobs Else savperobs = 0
If numdayspost > 0 Then savperday = totsav / numdayspost Else savperday = 0

'graph projection
ETMain.Graphbox.Cls
Dim grid, datapnts
color1 = BLACK
color2 = BLUE
grid = True
datapnts = True
Call TS2Graph(yfld, yfldproj, grid, datapnts)

'set energy type and output filename
If yfld = 4 Then
   etype = "Elec Use"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "kwh"
   eval = 1
ElseIf yfld = 5 Then
   etype = "Elec Demand"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "kwd"
   eval = 2
ElseIf yfld = 6 Then
   etype = "Fuel Use"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "thm"
   eval = 3
Else
   etype = "Energy"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "out"
End If

'print results
ETMain.StatusBox.Cls
Open outfilename$ For Output As #1

'print results to status box for 3P
ETMain.StatusBox.Print "Energy filename: "; efilename$
ETMain.StatusBox.Print "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%"
If numxvars = 1 Then
   ETMain.StatusBox.Print "Baseline model: "; etype; " = "; modeleqn
ElseIf numxvars = 2 Then
   ETMain.StatusBox.Print "Baseline model: "; etype; " = "; modeleqn; "  +  "; Format(ivcoefs(2), "#,##0.00"); " IV1"
ElseIf numxvars = 3 Then
   ETMain.StatusBox.Print "Baseline model: "; etype; " = "; modeleqn; "  +  "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
End If
            
'print results to output file
If numxvars = 1 Then
   Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00") '; "   IV1c = "; Format(ivcoefs(2), "#,##0.0000");"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
   Print #1, "Baseline model: "; etype; " = "; modeleqn
ElseIf numxvars = 2 Then
   Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000"); '"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
   Print #1, "Baseline model: "; etype; " = "; modeleqn; "  +  "; Format(ivcoefs(2), "#,##0.00"); " IV1"
ElseIf numxvars = 3 Then
   Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000"); "   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
   Print #1, "Baseline model: "; etype; " = "; modeleqn; "  +  "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
End If

'initialize global reg parameter variables
nxcp1 = -99
nxcp2 = -99
nycp = -99
nls = -99
nrs = -99

'save reg params as global variables for nac and multisite calculations
nmt = modtype
nxcp1 = xcp
nxcp2 = -99
nycp = ycp
nls = ls
nrs = rs
nivcoefs(2) = ivcoefs(2)
nivcoefs(3) = ivcoefs(3)

'fill coef matrix for drawing model lines on xy plots
For i = 1 To 3
   For j = 1 To 5
      c(i, j) = -99
   Next j
Next i
c(eval, 2) = xcp
c(eval, 3) = ycp
c(eval, 4) = ls
c(eval, 5) = rs

'print savings summary to screen
If totactpost > 0 And numpostobs > 0 And numdayspost > 0 Then
   ETMain.StatusBox.Print "Adjusted Baseline Use During Post-retrofit Period = "; Tab(60); Format(totadjbase, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   ETMain.StatusBox.Print "Actual Use During Post-retrofit Period = "; Tab(60); Format(totactpost, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   ETMain.StatusBox.Print "Total Savings During Post-retrofit Period = "; Tab(60); Format(totsav, "#,##0"); " ("; Format(percentsav, "0"); "%) +- "; Format(uncert, "#,##0.0"); " ("; Format(reluncert, "#,##0.0"); "%) over "; numpostobs; " observations and "; numdayspost; " days"
   ETMain.StatusBox.Print "Average Savings During Post-retrofit Period = "; Tab(60); Format(savperobs, "#,##0.00"); " per observation   =   "; Format(savperday, "#,##0.000"); " per day"
End If

'print savings summary to file
If totactpost > 0 And numpostobs > 0 And numdayspost > 0 Then
   Print #1, "Adjusted Baseline Use During Post-retrofit Period = "; Tab(60); Format(totadjbase, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   Print #1, "Actual Use During Post-retrofit Period = "; Tab(60); Format(totactpost, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   Print #1, "Total Savings During Post-retrofit Period = "; Tab(60); Format(totsav, "#,##0"); " ("; Format(percentsav, "0"); "%) +- "; Format(uncert, "#,##0.0"); " ("; Format(reluncert, "#,##0.0"); "%) over "; numpostobs; " observations and "; numdayspost; " days"
   Print #1, "Average Savings During Post-retrofit Period = "; Tab(60); Format(savperobs, "#,##0.00"); " per observation   =   "; Format(savperday, "#,##0.000"); " per day"
End If
Close #1

'print output data *.dat
outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "dat"
Open outfilename$ For Output As #1
For i = 1 To numrecs
   For j = 1 To 8
      Print #1, d(i, j),
   Next j
   For j = 9 To 12
      Print #1, Format(d(i, j), "#,##0.000"),
   Next j
   Print #1,
Next i
Close #1

'save reg params as global variables for nac calculations
nrmse = rmse
nn = n
ncvrmse = cvrmse
nr2 = r2

'save savings metrics as global variables for multilist output
nnumdayspost = numdayspost
nnumpostobs = numpostobs
ntotsav = totsav
nuncert = uncert
npercentsav = percentsav
nreluncert = reluncert
nsavperobs = savperobs
nsavperday = savperday
nmeanactpre = meanactpre
nmeanactpost = meanactpost

csce:
'display end message
Unload PerDone
Close #1
Screen.MousePointer = 0
'prints error message
If Err Then
   msg$ = """" + Error(Err) + """"
   MsgBox msg$, , "Error"
   Resume cses
End If
cses:

End Sub

Public Sub CalcSavings4PMVR(xfld, yfld, yfldproj, grpfld)
Dim i, j, numvars, numxvars
Dim rmse, cvrmse, r2, adjr2, roe, uncert, reluncert
Dim totsav, startdate, enddate, numdayspost, numpostobs
Dim xcp, cp1, cp2, ycp, ls, rs
Dim rmse4, rmse5, cvrmse4, cvrmse5, r24, r25
Dim modtype, etype, eval
Dim outfilename$
Dim totadjbase, totactpost
Dim percentsav, savperobs, savperday
Dim ytot, meanactpre, meanactpost
ReDim nivcoefs(8)

On Error GoTo csce
Screen.MousePointer = 11

'check if weather file is loaded
If numrecs = Empty Then
   msg$ = "Must open both energy and weather files before calculating savings."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf eopen <> True Then
   msg$ = "Must open energy file before calculating savings."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf wopen <> True Then
   msg$ = "Must open weather file before calculating savings."
   MsgBox msg$, , "Error"
   GoTo csce:
End If

'numindvars determined in "open energy cmd" after reading utl file
numxvars = numindvars

'dimension arrays and variables for calling cp-mvr models
ReDim x(numrecs, numxvars + 1), y(numrecs, numxvars + 1)
ReDim coef(numxvars + 1, 1), sec(numxvars + 1), sigma(numxvars + 1)
'ReDim ivcoefs(numxvars + 1), seivcoefs(numxvars + 1) 'note for cp-mvr models, coef for first ind var = ivcoefs(2), coef for second ind var = ivcoefs(3), etc.
ReDim ivcoefs(8), seivcoefs(8) 'note for cp-mvr models, coef for first ind var = ivcoefs(2), coef for second ind var = ivcoefs(3), etc.
Dim sexcp, seycp, sels, sers

'fill n, x(), y()
n = 0
meanactpre = 0
For i = 1 To numrecs
   If d(i, xfld) <> -99 And d(i, yfld) <> -99 And d(i, grpfld) = 1 Then
      If numxvars = 1 Or (numxvars = 2 And d(i, 9) <> -99) Or (numxvars = 3 And d(i, 9) <> -99 And d(i, 10) <> -99) Then
         'fill n, x() and y()
         n = n + 1
         x(n, 1) = 1
         x(n, 2) = d(i, xfld)
         If numxvars = 2 Then
            x(n, 3) = d(i, 9)
         ElseIf numxvars = 3 Then
            x(n, 3) = d(i, 9)
            x(n, 4) = d(i, 10)
         End If
         y(n, 1) = d(i, yfld)
         ytot = ytot + y(n, 1)
         meanactpre = ytot / n
         
         'finds max and mins
         If n = 1 Then
            xmin = x(1, 2)
            xmax = x(1, 2)
            ymin = y(1, 1)
            ymax = y(1, 1)
         End If
         If x(n, 2) < xmin Then xmin = x(n, 2)
         If x(n, 2) > xmax Then xmax = x(n, 2)
         If y(n, 1) < ymin Then ymin = y(n, 1)
         If y(n, 1) > ymax Then ymax = y(n, 1)
      End If
   End If
Next i

'call 4P model
rmse = -99
cvrmse = -99
r2 = -99
Call FourPMVR(x(), y(), n, numxvars, xcp, sexcp, ycp, seycp, ls, sels, rs, sers, ivcoefs(), seivcoefs(), rmse4, cvrmse4, r24)
modeleqn = Format(ycp, "#,##0.00") + "  -  " + Format(ls, "#,##0.00") + " (" + Format(xcp, "#,##0.00") + " - T)+  +  " + Format(rs, "#,##0.00") + " (T - " + Format(xcp, "#,##0.00") + ")+"

'set inf params
cvrmse = cvrmse4
r2 = r24
rmse = rmse4
modtype = "4P"

'init counting variables
numdayspost = 0
numpostobs = 0
totadjbase = 0
totactpost = 0
meanactpost = 0
totsav = 0
   
'fill d( ) with projected baseline model
For i = 1 To numrecs
   If d(i, grpfld) = 1 Then 'pre
      d(i, yfldproj) = d(i, yfld)
      startdate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
   Else 'post
      enddate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
      If d(i, xfld) <> ndflag And d(i, yfld) <> ndflag Then
         If numxvars = 1 Or (numxvars = 2 And d(i, 9) <> -99) Or (numxvars = 3 And d(i, 9) <> -99 And d(i, 10) <> -99) Then
            
            'calc projected energy use
            If d(i, xfld) <= xcp Then
              indvar1 = 1
              indvar2 = 0
            Else
               indvar1 = 0
               indvar2 = 1
            End If
            If numxvars = 1 Then
               d(i, yfldproj) = ycp - ls * indvar1 * (xcp - d(i, xfld)) + rs * indvar2 * (d(i, xfld) - xcp) ' + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
            ElseIf numxvars = 2 Then
               d(i, yfldproj) = ycp - ls * indvar1 * (xcp - d(i, xfld)) + rs * indvar2 * (d(i, xfld) - xcp) + ivcoefs(2) * d(i, 9)  '+ ivcoefs(3) * d(i, 10)
            ElseIf numxvars = 3 Then
               d(i, yfldproj) = ycp - ls * indvar1 * (xcp - d(i, xfld)) + rs * indvar2 * (d(i, xfld) - xcp) + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
            End If
            
            'calc post days and savings
            numdayspost = numdayspost + (enddate - startdate)
            numpostobs = numpostobs + 1
            totadjbase = totadjbase + d(i, yfldproj)
            totactpost = totactpost + d(i, yfld)
            meanactpost = totactpost / numpostobs
            totsav = totsav + (d(i, yfldproj) - d(i, yfld))
         End If
      Else
         d(i, yfldproj) = ndflag
      End If
      startdate = enddate
   End If
Next i

'calc uncertainty
If n > 0 And totsav <> 0 Then
   uncert = 1.96 * rmse * ((1 + 2 / n) * numpostobs) ^ 0.5
   reluncert = Abs(uncert / totsav) * 100
Else
   uncert = -99
   reluncert = -99
   totsav = 0
End If

'calc some savings metrics
If totactpost > 0 Then percentsav = totsav / totactpost * 100 Else percentsav = 0
If numpostobs > 0 Then savperobs = totsav / numpostobs Else savperobs = 0
If numdayspost > 0 Then savperday = totsav / numdayspost Else savperday = 0

'graph projection
ETMain.Graphbox.Cls
Dim grid, datapnts
color1 = BLACK
color2 = BLUE
grid = True
datapnts = True
Call TS2Graph(yfld, yfldproj, grid, datapnts)

'set energy type and output filename
If yfld = 4 Then
   etype = "Elec Use"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "kwh"
   eval = 1
ElseIf yfld = 5 Then
   etype = "Elec Demand"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "kwd"
   eval = 2
ElseIf yfld = 6 Then
   etype = "Fuel Use"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "thm"
   eval = 3
Else
   etype = "Energy"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "out"
End If

'print results
ETMain.StatusBox.Cls
Open outfilename$ For Output As #1


'print results to status box
ETMain.StatusBox.Print "Energy filename: "; efilename$
ETMain.StatusBox.Print "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00") '; "   IV1c = "; Format(ivcoefs(2), "#,##0.0000");"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
If numxvars = 1 Then
   ETMain.StatusBox.Print "Baseline model: "; etype; " = "; modeleqn
ElseIf numxvars = 2 Then
   ETMain.StatusBox.Print "Baseline model: "; etype; " = "; modeleqn; Format(ivcoefs(2), "#,##0.00"); " IV1"
ElseIf numxvars = 3 Then
   ETMain.StatusBox.Print "Baseline model: "; etype; " = "; modeleqn; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
End If
            
'print results to output file
If numxvars = 1 Then
   Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00") '; "   IV1c = "; Format(ivcoefs(2), "#,##0.0000");"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
   Print #1, "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(xcp, "#,##0.00"); ")+"
ElseIf numxvars = 2 Then
   Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000"); '"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
   Print #1, "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(xcp, "#,##0.00"); ")+  + "; Format(ivcoefs(2), "#,##0.00"); " IV1"
ElseIf numxvars = 3 Then
   Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000"); "   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
   Print #1, "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(xcp, "#,##0.00"); ")+  + "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
End If


'initialize global reg parameter variables
nxcp1 = -99
nxcp2 = -99
nycp = -99
nls = -99
nrs = -99

'save reg params as global variables for nac and multisite calcs
nmt = modtype
nxcp1 = xcp
nxcp2 = -99
nycp = ycp
nls = ls
nrs = rs
nivcoefs(2) = ivcoefs(2)
nivcoefs(3) = ivcoefs(3)

'fill coef matrix for drawing model lines on xy plots
For i = 1 To 3
   For j = 1 To 5
      c(i, j) = -99
   Next j
Next i
c(eval, 2) = xcp
c(eval, 3) = ycp
c(eval, 4) = ls
c(eval, 5) = rs
   


'print savings summary to screen
If totactpost > 0 And numpostobs > 0 And numdayspost > 0 Then
   ETMain.StatusBox.Print "Adjusted Baseline Use During Post-retrofit Period = "; Tab(60); Format(totadjbase, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   ETMain.StatusBox.Print "Actual Use During Post-retrofit Period = "; Tab(60); Format(totactpost, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   ETMain.StatusBox.Print "Total Savings During Post-retrofit Period = "; Tab(60); Format(totsav, "#,##0"); " ("; Format(percentsav, "0"); "%) +- "; Format(uncert, "#,##0.0"); " ("; Format(reluncert, "#,##0.0"); "%) over "; numpostobs; " observations and "; numdayspost; " days"
   ETMain.StatusBox.Print "Average Savings During Post-retrofit Period = "; Tab(60); Format(savperobs, "#,##0.00"); " per observation   =   "; Format(savperday, "#,##0.000"); " per day"
End If

'print savings summary to file
If totactpost > 0 And numpostobs > 0 And numdayspost > 0 Then
   Print #1, "Adjusted Baseline Use During Post-retrofit Period = "; Tab(60); Format(totadjbase, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   Print #1, "Actual Use During Post-retrofit Period = "; Tab(60); Format(totactpost, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   Print #1, "Total Savings During Post-retrofit Period = "; Tab(60); Format(totsav, "#,##0"); " ("; Format(percentsav, "0"); "%) +- "; Format(uncert, "#,##0.0"); " ("; Format(reluncert, "#,##0.0"); "%) over "; numpostobs; " observations and "; numdayspost; " days"
   Print #1, "Average Savings During Post-retrofit Period = "; Tab(60); Format(savperobs, "#,##0.00"); " per observation   =   "; Format(savperday, "#,##0.000"); " per day"
End If
Close #1

'print output data *.dat
outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "dat"
Open outfilename$ For Output As #1
For i = 1 To numrecs
   For j = 1 To 8
      Print #1, d(i, j),
   Next j
   For j = 9 To 12
      Print #1, Format(d(i, j), "#,##0.000"),
   Next j
   Print #1,
Next i
Close #1

'save reg params as global variables for nac calculations
nrmse = rmse
nn = n
ncvrmse = cvrmse
nr2 = r2

'save savings metrics as global variables for multilist output
nnumdayspost = numdayspost
nnumpostobs = numpostobs
ntotsav = totsav
nuncert = uncert
npercentsav = percentsav
nreluncert = reluncert
nsavperobs = savperobs
nsavperday = savperday
nmeanactpre = meanactpre
nmeanactpost = meanactpost

csce:
'display end message
Unload PerDone
Close #1
Screen.MousePointer = 0
'prints error message
If Err Then
   msg$ = """" + Error(Err) + """"
   MsgBox msg$, , "Error"
   Resume cses
End If
cses:

End Sub

Public Sub CalcSavings5PMVR(xfld, yfld, yfldproj, grpfld)
Dim i, j, numvars, numxvars
Dim rmse, cvrmse, r2, adjr2, roe, uncert, reluncert
Dim totsav, startdate, enddate, numdayspost, numpostobs
Dim xcp, cp1, cp2, ycp, ls, rs
Dim rmse4, rmse5, cvrmse4, cvrmse5, r24, r25
Dim modtype, etype, eval
Dim outfilename$
Dim totadjbase, totactpost
Dim percentsav, savperobs, savperday
Dim ytot, meanactpre, meanactpost
ReDim nivcoefs(8)

On Error GoTo csce
Screen.MousePointer = 11

'check if weather file is loaded
If numrecs = Empty Then
   msg$ = "Must open both energy and weather files before calculating savings."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf eopen <> True Then
   msg$ = "Must open energy file before calculating savings."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf wopen <> True Then
   msg$ = "Must open weather file before calculating savings."
   MsgBox msg$, , "Error"
   GoTo csce:
End If

'numindvars determined in "open energy cmd" after reading utl file
numxvars = numindvars

'dimension arrays and variables for calling cp-mvr models
ReDim x(numrecs, numxvars + 1), y(numrecs, numxvars + 1)
ReDim coef(numxvars + 1, 1), sec(numxvars + 1), sigma(numxvars + 1)
'ReDim ivcoefs(numxvars + 1), seivcoefs(numxvars + 1) 'note for cp-mvr models, coef for first ind var = ivcoefs(2), coef for second ind var = ivcoefs(3), etc.
ReDim ivcoefs(8), seivcoefs(8) 'note for cp-mvr models, coef for first ind var = ivcoefs(2), coef for second ind var = ivcoefs(3), etc.
Dim sexcp, seycp, sels, sers

'fill n, x(), y()
n = 0
meanactpre = 0
For i = 1 To numrecs
   If d(i, xfld) <> -99 And d(i, yfld) <> -99 And d(i, grpfld) = 1 Then
      If numxvars = 1 Or (numxvars = 2 And d(i, 9) <> -99) Or (numxvars = 3 And d(i, 9) <> -99 And d(i, 10) <> -99) Then
         'fill n, x() and y()
         n = n + 1
         x(n, 1) = 1
         x(n, 2) = d(i, xfld)
         If numxvars = 2 Then
            x(n, 3) = d(i, 9)
         ElseIf numxvars = 3 Then
            x(n, 3) = d(i, 9)
            x(n, 4) = d(i, 10)
         End If
         y(n, 1) = d(i, yfld)
         ytot = ytot + y(n, 1)
         meanactpre = ytot / n
   
         
         'finds max and mins
         If n = 1 Then
            xmin = x(1, 2)
            xmax = x(1, 2)
            ymin = y(1, 1)
            ymax = y(1, 1)
         End If
         If x(n, 2) < xmin Then xmin = x(n, 2)
         If x(n, 2) > xmax Then xmax = x(n, 2)
         If y(n, 1) < ymin Then ymin = y(n, 1)
         If y(n, 1) > ymax Then ymax = y(n, 1)
      End If
   End If
Next i


'select 5p model
rmse = -99
cvrmse = -99
r2 = -99
Call FivePMVR(x(), y(), n, numxvars, cp1, sexcp1, cp2, sexcp2, ycp, seycp, ls, sels, rs, sers, ivcoefs(), seivcoefs(), rmse5, cvrmse5, r25)
modeleqn = Format(ycp, "#,##0.00") + "  -  " + Format(ls, "#,##0.00") + " (" + Format(cp1, "#,##0.00") + " - T)+  +  " + Format(rs, "#,##0.00") + " (T - " + Format(cp2, "#,##0.00") + ")+"

'set inf params
modtype = "5P"
cvrmse = cvrmse5
rmse = rmse5
r2 = r25

'init counting variables
numdayspost = 0
numpostobs = 0
totadjbase = 0
totactpost = 0
meanactpost = 0
totsav = 0

'fill d() with projected baseline model
For i = 1 To numrecs
   If d(i, grpfld) = 1 Then 'pre
      d(i, yfldproj) = d(i, yfld)
      startdate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
   Else 'post
      enddate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
      If d(i, xfld) <> ndflag And d(i, yfld) <> ndflag Then
         If numxvars = 1 Or (numxvars = 2 And d(i, 9) <> -99) Or (numxvars = 3 And d(i, 9) <> -99 And d(i, 10) <> -99) Then

         
            'calc projected energy use
            If d(i, xfld) <= cp1 Then
              indvar1 = 1
            Else
               indvar1 = 0
            End If
            If d(i, xfld) <= cp2 Then
               indvar2 = 0
            Else
               indvar2 = 1
            End If
            If numxvars = 1 Then
               d(i, yfldproj) = ycp + ls * indvar1 * (d(i, xfld) - cp1) + rs * indvar2 * (d(i, xfld) - cp2) ' + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
            ElseIf numxvars = 2 Then
               d(i, yfldproj) = ycp + ls * indvar1 * (d(i, xfld) - cp1) + rs * indvar2 * (d(i, xfld) - cp2) + ivcoefs(2) * d(i, 9) '+ ivcoefs(3) * d(i, 10)
            ElseIf numxvars = 3 Then
               d(i, yfldproj) = ycp + ls * indvar1 * (d(i, xfld) - cp1) + rs * indvar2 * (d(i, xfld) - cp2) + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
            End If
            
            'calc post days and savings
            numdayspost = numdayspost + (enddate - startdate)
            numpostobs = numpostobs + 1
            totadjbase = totadjbase + d(i, yfldproj)
            totactpost = totactpost + d(i, yfld)
            meanactpost = totactpost / numpostobs
            totsav = totsav + (d(i, yfldproj) - d(i, yfld))
         End If
      Else
         d(i, yfldproj) = ndflag
      End If
      startdate = enddate
   End If
Next i


'calc uncertainty
If n > 0 And totsav <> 0 Then
   uncert = 1.96 * rmse * ((1 + 2 / n) * numpostobs) ^ 0.5
   reluncert = Abs(uncert / totsav) * 100
Else
   uncert = -99
   reluncert = -99
   totsav = 0
End If

'calc some savings metrics
If totactpost > 0 Then percentsav = totsav / totactpost * 100 Else percentsav = 0
If numpostobs > 0 Then savperobs = totsav / numpostobs Else savperobs = 0
If numdayspost > 0 Then savperday = totsav / numdayspost Else savperday = 0

'graph projection
ETMain.Graphbox.Cls
Dim grid, datapnts
color1 = BLACK
color2 = BLUE
grid = True
datapnts = True
Call TS2Graph(yfld, yfldproj, grid, datapnts)

'set energy type and output filename
If yfld = 4 Then
   etype = "Elec Use"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "kwh"
   eval = 1
ElseIf yfld = 5 Then
   etype = "Elec Demand"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "kwd"
   eval = 2
ElseIf yfld = 6 Then
   etype = "Fuel Use"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "thm"
   eval = 3
Else
   etype = "Energy"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "out"
End If

'print results
ETMain.StatusBox.Cls
Open outfilename$ For Output As #1
   
'print results to status box
ETMain.StatusBox.Print "Energy filename: "; efilename$
ETMain.StatusBox.Print "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%"
If numxvars = 1 Then
   ETMain.StatusBox.Print "Baseline model: "; etype; " = "; modeleqn
ElseIf numxvars = 2 Then
   ETMain.StatusBox.Print "Baseline model: "; etype; " = "; modeleqn; " + "; Format(ivcoefs(2), "#,##0.00"); " IV1"
ElseIf numxvars = 3 Then
   ETMain.StatusBox.Print "Baseline model: "; etype; " = "; modeleqn; "  + "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
End If

'print results to output file
If numxvars = 1 Then
   Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp1 = "; Format(cp1, "#,##0.00"); "   Xcp2 = "; Format(cp2, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00") '; "   IV1c = "; Format(ivcoefs(2), "#,##0.0000");"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
   Print #1, "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(cp1, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(cp2, "#,##0.00"); ")+" ' + "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
ElseIf numxvars = 2 Then
   Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp1 = "; Format(cp1, "#,##0.00"); "   Xcp2 = "; Format(cp2, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000") ';"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
   Print #1, "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(cp1, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(cp2, "#,##0.00"); ")+  + "; Format(ivcoefs(2), "#,##0.00"); " IV1" ';  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
ElseIf numxvars = 3 Then
   Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp1 = "; Format(cp1, "#,##0.00"); "   Xcp2 = "; Format(cp2, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000"); "   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
   Print #1, "Baseline model: "; etype; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(cp1, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(cp2, "#,##0.00"); ")+  + "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
End If

'initialize global reg parameter variables
nxcp1 = -99
nxcp2 = -99
nycp = -99
nls = -99
nrs = -99


'save reg params as global variables for nac calc
nmt = modtype
nxcp1 = cp1
nxcp2 = cp2
nycp = ycp
nls = ls
nrs = rs
nivcoefs(2) = ivcoefs(2)
nivcoefs(3) = ivcoefs(3)

'fill coef matrix for drawing model lines on xy plots
For i = 1 To 3
   For j = 1 To 5
      c(i, j) = -99
   Next j
Next i
c(eval, 1) = cp1
c(eval, 2) = cp2
c(eval, 3) = ycp
c(eval, 4) = ls
c(eval, 5) = rs

'print savings summary to screen
If totactpost > 0 And numpostobs > 0 And numdayspost > 0 Then
   ETMain.StatusBox.Print "Adjusted Baseline Use During Post-retrofit Period = "; Tab(60); Format(totadjbase, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   ETMain.StatusBox.Print "Actual Use During Post-retrofit Period = "; Tab(60); Format(totactpost, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   ETMain.StatusBox.Print "Total Savings During Post-retrofit Period = "; Tab(60); Format(totsav, "#,##0"); " ("; Format(percentsav, "0"); "%) +- "; Format(uncert, "#,##0.0"); " ("; Format(reluncert, "#,##0.0"); "%) over "; numpostobs; " observations and "; numdayspost; " days"
   ETMain.StatusBox.Print "Average Savings During Post-retrofit Period = "; Tab(60); Format(savperobs, "#,##0.00"); " per observation   =   "; Format(savperday, "#,##0.000"); " per day"
End If

'print savings summary to file
If totactpost > 0 And numpostobs > 0 And numdayspost > 0 Then
   Print #1, "Adjusted Baseline Use During Post-retrofit Period = "; Tab(60); Format(totadjbase, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   Print #1, "Actual Use During Post-retrofit Period = "; Tab(60); Format(totactpost, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   Print #1, "Total Savings During Post-retrofit Period = "; Tab(60); Format(totsav, "#,##0"); " ("; Format(percentsav, "0"); "%) +- "; Format(uncert, "#,##0.0"); " ("; Format(reluncert, "#,##0.0"); "%) over "; numpostobs; " observations and "; numdayspost; " days"
   Print #1, "Average Savings During Post-retrofit Period = "; Tab(60); Format(savperobs, "#,##0.00"); " per observation   =   "; Format(savperday, "#,##0.000"); " per day"
End If
Close #1

'print output data *.dat
outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "dat"
Open outfilename$ For Output As #1
For i = 1 To numrecs
   For j = 1 To 8
      Print #1, d(i, j),
   Next j
   For j = 9 To 12
      Print #1, Format(d(i, j), "#,##0.000"),
   Next j
   Print #1,
Next i
Close #1

'save reg params as global variables for nac calculations
nrmse = rmse
nn = n
ncvrmse = cvrmse
nr2 = r2

'save savings metrics as global variables for multilist output
nnumdayspost = numdayspost
nnumpostobs = numpostobs
ntotsav = totsav
nuncert = uncert
npercentsav = percentsav
nreluncert = reluncert
nsavperobs = savperobs
nsavperday = savperday
nmeanactpre = meanactpre
nmeanactpost = meanactpost

csce:
'display end message
Unload PerDone
Close #1
Screen.MousePointer = 0
'prints error message
If Err Then
   msg$ = """" + Error(Err) + """"
   MsgBox msg$, , "Error"
   Resume cses
End If
cses:
End Sub




Public Sub TwoPMVR(x(), y(), n, numxvars, aa, seaa, bb, sebb, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)
Dim i, j, k, xmin, xmax, inc, numints
Dim itration, xmin4p, n1, n2, cp
Dim IndVar, rmsemin, bestcp, t, predint
Dim betar, betac
Dim totregs, regnum 'for perdone

'turns on VB error handling
On Error GoTo cpce

'show perdone
Screen.MousePointer = 11
PerDone.Caption = "Processing Data..."
PerDone.Show 0
DoEvents

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo cpce:
End If

'set modtype (global variable)
modtype = "2PMVR"

'dimension arrays
'numxvars = 1
ReDim A(n, numxvars + 2), beta(numxvars + 2, 1), sec(numxvars + 2), sigma(numxvars + 2)


'fills A()
'cp = xmin4p + i * inc
'For j = 1 To n
'   If x(j, 2) <= cp Then
'      n1 = n1 + 1
'      IndVar = 0
'   Else
'      n2 = n2 + 1
'      IndVar = 1
'   End If
'   A(j, 1) = 1
'   A(j, 2) = x(j, 2)
'   A(j, 3) = IndVar * (x(j, 2) - cp)
'
'   For k = 1 To numxvars - 1
'      A(j, 3 + k) = x(j, 2 + k)
'   Next k
'Next j

'calls regression engine
'Call Reg(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec())
'Call Inf(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec(), rmse, cvrmse, r2, adjr2, sigma(), roe)
Call Reg(x(), n, numxvars + 1, y(), n, 1, beta(), betar, betac, sec())
Call Inf(x(), n, numxvars + 1, y(), n, 1, beta(), betar, betac, sec(), rmse, cvrmse, r2, adjr2, sigma(), roe)
'for 2P model form y = aa + bb T
aa = beta(1, 1)
bb = beta(2, 1)
seaa = sigma(1)
sebb = sigma(2)
For k = 2 To numxvars
   ivcoefs(k) = beta(k + 1, 1)
   seivcoefs(k) = sigma(k + 1)
Next k

'calcs prediction interval
't = 1.96 + 2.7 / (n - 3)
'predint = t * rmse * (1 + 1 / n) ^ 0.5

'vb error handling
cpce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   MsgBox """" + Error(Err) + """"
   Resume cpes
End If
cpes:
End Sub

Public Sub CalcSavings2PMVR(xfld, yfld, yfldproj, grpfld)
Dim i, j, numvars, numxvars
Dim rmse, cvrmse, r2, adjr2, roe, uncert, reluncert
Dim totsav, startdate, enddate, numdayspost, numpostobs
Dim xcp, cp1, cp2, ycp, ls, rs
Dim rmse4, rmse5, cvrmse4, cvrmse5, r24
Dim modtype, etype, eval
Dim outfilename$
Dim totadjbase, totactpost
Dim percentsav, savperobs, savperday
Dim ytot, meanactpre, meanactpost
ReDim nivcoefs(8)

On Error GoTo csce
Screen.MousePointer = 11

'check if weather file is loaded
If numrecs = Empty Then
   msg$ = "Must open both energy and weather files before calculating savings."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf eopen <> True Then
   msg$ = "Must open energy file before calculating savings."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf wopen <> True Then
   msg$ = "Must open weather file before calculating savings."
   MsgBox msg$, , "Error"
   GoTo csce:
End If

'numindvars determined in "open energy cmd" after reading utl file
numxvars = numindvars

'dimension arrays and variables for calling models
ReDim x(numrecs, numxvars + 1), y(numrecs, numxvars + 1)
ReDim coef(numxvars + 1, 1), sec(numxvars + 1), sigma(numxvars + 1)
'ReDim ivcoefs(numxvars + 1), seivcoefs(numxvars + 1) 'note for cp-mvr models, coef for first ind var = ivcoefs(2), coef for second ind var = ivcoefs(3), etc.
ReDim ivcoefs(8), seivcoefs(8) 'note for cp-mvr models, coef for first ind var = ivcoefs(2), coef for second ind var = ivcoefs(3), etc.
Dim sexcp, seycp, sels, sers

'fill n, x(), y()
n = 0
meanactpre = 0
For i = 1 To numrecs
   If d(i, xfld) <> -99 And d(i, yfld) <> -99 And d(i, grpfld) = 1 Then
      If numxvars = 1 Or (numxvars = 2 And d(i, 9) <> -99) Or (numxvars = 3 And d(i, 9) <> -99 And d(i, 10) <> -99) Then
         'fill n, x() and y()
         n = n + 1
         x(n, 1) = 1
         x(n, 2) = d(i, xfld)
         If numxvars = 2 Then
            x(n, 3) = d(i, 9)
         ElseIf numxvars = 3 Then
            x(n, 3) = d(i, 9)
            x(n, 4) = d(i, 10)
         End If
         y(n, 1) = d(i, yfld)
         ytot = ytot + y(n, 1)
         meanactpre = ytot / n
         
         'finds max and mins
         If n = 1 Then
            xmin = x(1, 2)
            xmax = x(1, 2)
            ymin = y(1, 1)
            ymax = y(1, 1)
         End If
         If x(n, 2) < xmin Then xmin = x(n, 2)
         If x(n, 2) > xmax Then xmax = x(n, 2)
         If y(n, 1) < ymin Then ymin = y(n, 1)
         If y(n, 1) > ymax Then ymax = y(n, 1)
      End If
   End If
Next i

'call 2P model
rmse = -99
cvrmse = -99
r2 = -99
Call TwoPMVR(x(), y(), n, numxvars, aa, seaa, bb, sebb, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)
modeleqn = Format(aa, "#,##0.00") + "  +  " + Format(bb, "#,##0.00") + " T"

'set inf params
modtype = "2P"

'init counting variables
numdayspost = 0
numpostobs = 0
totadjbase = 0
totactpost = 0
meanactpost = 0
totsav = 0
   
'fill d( ) with projected baseline model
For i = 1 To numrecs
   If d(i, grpfld) = 1 Then 'pre
      d(i, yfldproj) = d(i, yfld)
      startdate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
   Else 'post
      enddate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
      If d(i, xfld) <> ndflag And d(i, yfld) <> ndflag Then
         If numxvars = 1 Or (numxvars = 2 And d(i, 9) <> -99) Or (numxvars = 3 And d(i, 9) <> -99 And d(i, 10) <> -99) Then
               
            'calc projected energy use
            If numxvars = 1 Then
               d(i, yfldproj) = aa + bb * d(i, xfld)
            ElseIf numxvars = 2 Then
               d(i, yfldproj) = aa + bb * d(i, xfld) + ivcoefs(2) * d(i, 9)
            ElseIf numxvars = 3 Then
               d(i, yfldproj) = aa + bb * d(i, xfld) + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
            End If
            
            'calc post days and savings
            numdayspost = numdayspost + (enddate - startdate)
            numpostobs = numpostobs + 1
            totadjbase = totadjbase + d(i, yfldproj)
            totactpost = totactpost + d(i, yfld)
            meanactpost = totactpost / numpostobs
            totsav = totsav + (d(i, yfldproj) - d(i, yfld))
         End If
      Else
         d(i, yfldproj) = ndflag
      End If
      startdate = enddate
   End If
Next i

'calc uncertainty
If n > 0 And totsav <> 0 Then
   uncert = 1.96 * rmse * ((1 + 2 / n) * numpostobs) ^ 0.5
   reluncert = Abs(uncert / totsav) * 100
Else
   uncert = -99
   reluncert = -99
   totsav = 0
End If

'calc some savings metrics
If totactpost > 0 Then percentsav = totsav / totactpost * 100 Else percentsav = 0
If numpostobs > 0 Then savperobs = totsav / numpostobs Else savperobs = 0
If numdayspost > 0 Then savperday = totsav / numdayspost Else savperday = 0

'graph projection
ETMain.Graphbox.Cls
Dim grid, datapnts
color1 = BLACK
color2 = BLUE
grid = True
datapnts = True
Call TS2Graph(yfld, yfldproj, grid, datapnts)

'set energy type and output filename
If yfld = 4 Then
   etype = "Elec Use"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "kwh"
   eval = 1
ElseIf yfld = 5 Then
   etype = "Elec Demand"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "kwd"
   eval = 2
ElseIf yfld = 6 Then
   etype = "Fuel Use"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "thm"
   eval = 3
Else
   etype = "Energy"
   outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "out"
End If

'print results
ETMain.StatusBox.Cls
Open outfilename$ For Output As #1

'print results to status box
ETMain.StatusBox.Print "Energy filename: "; efilename$
ETMain.StatusBox.Print "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00") '; "   IV1c = "; Format(ivcoefs(2), "#,##0.0000");"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
If numxvars = 1 Then
   ETMain.StatusBox.Print "Baseline model: "; etype; " = "; modeleqn
ElseIf numxvars = 2 Then
   ETMain.StatusBox.Print "Baseline model: "; etype; " = "; modeleqn; "  +  "; Format(ivcoefs(2), "#,##0.00"); " IV1"
ElseIf numxvars = 3 Then
   ETMain.StatusBox.Print "Baseline model: "; etype; " = "; modeleqn; "  +  "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
End If
            
'print results to output file
If numxvars = 1 Then
   Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00") '; "   IV1c = "; Format(ivcoefs(2), "#,##0.0000");"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
   Print #1, "Baseline model: "; etype; " = "; modeleqn
ElseIf numxvars = 2 Then
   Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000"); '"   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
   Print #1, "Baseline model: "; etype; " = "; modeleqn; "  +  "; Format(ivcoefs(2), "#,##0.00"); " IV1"
ElseIf numxvars = 3 Then
   Print #1, "Baseline model stats: "; modtype; "   N = "; n; "   R2 = "; Format(r2, "0.00"); "   CV-RMSE = "; Format(cvrmse, "0.0"); "%" '   Xcp = "; Format(xcp, "#,##0.00"); "   Ycp = "; Format(ycp, "#,##0.00"); "   LS = "; Format(ls, "#,##0.00"); "   RS = "; Format(rs, "#,##0.00"); "   IV1c = "; Format(ivcoefs(2), "#,##0.0000"); "   IV2c = "; Format(ivcoefs(3), "#,##0.0000")
   Print #1, "Baseline model: "; etype; " = "; modeleqn; "  +  "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
End If
   
'print savings summary to screen
If totactpost > 0 And numpostobs > 0 And numdayspost > 0 Then
   ETMain.StatusBox.Print "Adjusted Baseline Use During Post-retrofit Period = "; Tab(60); Format(totadjbase, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   ETMain.StatusBox.Print "Actual Use During Post-retrofit Period = "; Tab(60); Format(totactpost, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   ETMain.StatusBox.Print "Total Savings During Post-retrofit Period = "; Tab(60); Format(totsav, "#,##0"); " ("; Format(percentsav, "0"); "%) +- "; Format(uncert, "#,##0.0"); " ("; Format(reluncert, "#,##0.0"); "%) over "; numpostobs; " observations and "; numdayspost; " days"
   ETMain.StatusBox.Print "Average Savings During Post-retrofit Period = "; Tab(60); Format(savperobs, "#,##0.00"); " per observation   =   "; Format(savperday, "#,##0.000"); " per day"
End If

'print savings summary to file
If totactpost > 0 And numpostobs > 0 And numdayspost > 0 Then
   Print #1, "Adjusted Baseline Use During Post-retrofit Period = "; Tab(60); Format(totadjbase, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   Print #1, "Actual Use During Post-retrofit Period = "; Tab(60); Format(totactpost, "#,##0"); " over "; numpostobs; " observations and "; numdayspost; " days"
   Print #1, "Total Savings During Post-retrofit Period = "; Tab(60); Format(totsav, "#,##0"); " ("; Format(percentsav, "0"); "%) +- "; Format(uncert, "#,##0.0"); " ("; Format(reluncert, "#,##0.0"); "%) over "; numpostobs; " observations and "; numdayspost; " days"
   Print #1, "Average Savings During Post-retrofit Period = "; Tab(60); Format(savperobs, "#,##0.00"); " per observation   =   "; Format(savperday, "#,##0.000"); " per day"
End If
Close #1

'print output data *.dat
outfilename$ = Mid$(efilepath$ + efilename$, 1, Len(efilepath$ + efilename$) - 3) + "dat"
Open outfilename$ For Output As #1
For i = 1 To numrecs
   For j = 1 To 8
      Print #1, d(i, j),
   Next j
   For j = 9 To 12
      Print #1, Format(d(i, j), "#,##0.000"),
   Next j
   Print #1,
Next i
Close #1

'initialize global reg parameter variables
nxcp1 = -99
nxcp2 = -99
nycp = -99
nls = -99
nrs = -99

'save reg params as global variables for nac calculations and multisite output
nmt = modtype
nycp = aa  'for 2p model of form y = aa + bbT, place aa in ycp field and bb in ls field of multisite output file
nls = bb   'for 2p model of form y = aa + bbT, place aa in ycp field and bb in ls field of multisite output file
nivcoefs(2) = ivcoefs(2)
nivcoefs(3) = ivcoefs(3)
nrmse = rmse
nn = n
ncvrmse = cvrmse
nr2 = r2

'fill coef matrix for drawing model lines on xy plots
For i = 1 To 3
   For j = 1 To 5
      c(i, j) = -99
   Next j
Next i
c(eval, 3) = aa 'ycp
c(eval, 4) = bb 'ls

'save savings metrics as global variables for multilist output
nnumdayspost = numdayspost
nnumpostobs = numpostobs
ntotsav = totsav
nuncert = uncert
npercentsav = percentsav
nreluncert = reluncert
nsavperobs = savperobs
nsavperday = savperday
nmeanactpre = meanactpre
nmeanactpost = meanactpost

csce:
'display end message
Unload PerDone
Close #1
Screen.MousePointer = 0
'prints error message
If Err Then
   msg$ = """" + Error(Err) + """"
   MsgBox msg$, , "Error"
   Resume cses
End If
cses:

End Sub

Public Sub CalcSSSliding(xfld, yfld, yfldproj, grpfld)
Dim i, j
Dim startdate, enddate, numdayspost, numpostobs
Dim xcp, ycp, ls, rs, cp1, cp2
ReDim ivcoefs(8)
Dim numdays, ttot, ttot2, wdate, numxvars
ReDim toanorm(numrecs), tiv1norm(numrecs), tiv2norm(numrecs)
ReDim nac(numrecs), ac(numrecs), nxcp1s(numrecs), nxcp2s(numrecs), nycps(numrecs), nlss(numrecs), nrss(numrecs), nivcoef1s(numrecs), nivcoef2s(numrecs)
ReDim nivcoefs(8)
Dim numnacs, NACTot, TbalTot, EindTot, STot

On Error GoTo csce
Screen.MousePointer = 11

'check if weather, energy and tm2 files are loaded
If numrecs = Empty Then
   msg$ = "Must open both energy and weather files before calculating sliding NAC."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf eopen <> True Then
   msg$ = "Must open energy file before calculating Sliding NAC."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf wopen <> True Then
   msg$ = "Must open weather file before calculating sliding NAC."
   MsgBox msg$, , "Error"
   GoTo csce:
ElseIf tm2open <> True Then
   msg$ = "Must open TMY2 weather data file before calculating sliding NAC."
   MsgBox msg$, , "Error"
   GoTo csce:
End If

If msslidinganal <> True Then 'single site
   'open sliding output file
   If engytype$ = "elec" Then
      sloutfile$ = efilepath$ + Mid$(efilename$, 1, Len(efilename$) - 4) + "_SlidingElecOutput.txt"
   ElseIf engytype$ = "fuel" Then
      sloutfile$ = efilepath$ + Mid$(efilename$, 1, Len(efilename$) - 4) + "_SlidingFuelOutput.txt"
   Else
      Stop
   End If
   Open sloutfile$ For Output As #9
   'print sliding output file header line
   mstats = "Sitename" + Chr(9) + "ModelType" + Chr(9) + "End Date" + Chr(9) + "Npre" + Chr(9) + "Meanpre" + Chr(9) + "R2" + Chr(9) + "RMSE" + Chr(9) + "CV-RMSE"
   mcoefs = "xcp1" + Chr(9) + "xcp2" + Chr(9) + "ycp" + Chr(9) + "ls" + Chr(9) + "rs" + Chr(9) + "iv1" + Chr(9) + "iv2"
   savsum = "AC" + Chr(9) + "NAC" ' + Chr(9) + "Npst" + Chr(9) + "Meanpst" + Chr(9) + "Dayspst" + Chr(9) + "Sav" + Chr(9) + "e_Sav" + Chr(9) + "PerSav" + Chr(9) + "e_PerSav" + Chr(9) + "SavPerObs" + Chr(9) + "SavPerDay"
   Print #9, mstats + Chr(9) + mcoefs + Chr(9) + savsum
End If

'zero as() and nac()
For i = 1 To numrecs
   ac(i) = 0
   nac(i) = 0
Next i

'numindvars determined in "open energy cmd" after reading utl file
numxvars = numindvars

'fill Toanorm() from tmy2 data
startdate = CDate(Format(e(0, 1)) + "/" + Format(e(0, 2)) + "/" + Format(e(0, 3)))
For i = 1 To numrecs
   enddate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
   
   'calc normal temp during energy period from tmy2 avg daily temps
   If Year(startdate) = Year(enddate) Then
      numdays = 0
      ttot = 0
      For j = 1 To 365
         wdate = CDate(Format(w(j, 1)) + "/" + Format(w(j, 2)) + "/" + Format(Year(startdate)))
         If wdate > startdate And wdate <= enddate Then
            numdays = numdays + 1
            ttot = ttot + w(j, 4)
         End If
      Next j
      toanorm(i) = ttot / numdays
   Else 'startyear < endyear
      numdays = 0
      ttot = 0
      'calc sum of temps from startdate to end of calendar year
      For j = 1 To 365
         wdate = CDate(Format(w(j, 1)) + "/" + Format(w(j, 2)) + "/" + Format(Year(startdate)))
         If wdate > startdate Then
            numdays = numdays + 1
            ttot = ttot + w(j, 4)
         End If
      Next j
      'add sum of temps from start of calendar year to enddate
      For j = 1 To 365
         wdate = CDate(Format(w(j, 1)) + "/" + Format(w(j, 2)) + "/" + Format(Year(enddate)))
         If wdate < enddate Then
            numdays = numdays + 1
            ttot = ttot + w(j, 4)
         End If
      Next j
      'calc average temp during period
      toanorm(i) = ttot / numdays
   End If
   startdate = enddate
Next i
'for i = 1 to numrecs: print i, toanorm(i), d(i,11): next i

'if numxvars > 1 then fill tiv1norm() and tiv2norm() from tiv() where tiv(i,1) = mo, tiv(i,2) = dy, tiv(i,3) = typindvar1, tiv(i,4) = typindvar2
If numxvars > 1 Then
   
   'check if tivfile is open
   If tivopen <> True Then
      msg$ = "Energy file (*.utl) includes independent variables, but typical independent variable file (*.tiv) is not open.  If single site analysis, Open TIV file after opening TM2 file.  If multisite analysis, add TIV file to Multisite List File."
      MsgBox msg$, , "Error"
      GoTo csce
   End If
      
   startdate = CDate(Format(e(0, 1)) + "/" + Format(e(0, 2)) + "/" + Format(e(0, 3)))
   For i = 1 To numrecs
      enddate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
      
      'calc normal temp during energy period from tmy2 avg daily temps
      If Year(startdate) = Year(enddate) Then
         numdays = 0
         ttot = 0
         ttot2 = 0
         For j = 1 To 365
            wdate = CDate(Format(tiv(j, 1)) + "/" + Format(tiv(j, 2)) + "/" + Format(Year(startdate)))
            If wdate > startdate And wdate <= enddate Then
               numdays = numdays + 1
               ttot = ttot + tiv(j, 3)
               If numxvars = 3 Then ttot2 = ttot2 + tiv(j, 4)
            End If
         Next j
         tiv1norm(i) = ttot / numdays
         If numxvars = 3 Then tiv2norm(i) = ttot2 / numdays
         
      Else 'startyear < endyear
         numdays = 0
         ttot = 0
         ttot2 = 0
         'calc sum of temps from startdate to end of calendar year
         For j = 1 To 365
            wdate = CDate(Format(tiv(j, 1)) + "/" + Format(tiv(j, 2)) + "/" + Format(Year(startdate)))
            If wdate > startdate Then
               numdays = numdays + 1
               ttot = ttot + tiv(j, 3)
               If numxvars = 3 Then ttot2 = ttot2 + tiv(j, 4)
            End If
         Next j
         'add sum of temps from start of calendar year to enddate
         For j = 1 To 365
            wdate = CDate(Format(tiv(j, 1)) + "/" + Format(tiv(j, 2)) + "/" + Format(Year(enddate)))
            If wdate < enddate Then
               numdays = numdays + 1
               ttot = ttot + tiv(j, 3)
               If numxvars = 3 Then ttot2 = ttot2 + tiv(j, 4)
            End If
         Next j
         'calc average temp during period
         tiv1norm(i) = ttot / numdays
         If numxvars = 3 Then tiv2norm(i) = ttot2 / numdays
      End If
      startdate = enddate
   Next i
End If
'test in immediate window For i = 1 To numrecs: Print i, tiv1norm(i), d(i, 9), tiv2norm(i), d(i, 10): Next i


'dimension arrays and variables for calling models
ReDim x(numrecs, numxvars + 1), y(numrecs, numxvars + 1)
ReDim coef(numxvars + 1, 1), sec(numxvars + 1), sigma(numxvars + 1)
ReDim ivcoefs(8), seivcoefs(8) 'note for cp-mvr models, coef for first ind var = ivcoefs(2), coef for second ind var = ivcoefs(3), etc.
Dim sexcp, seycp, sels, sers

'calc numsets of sliding regressions and nac calculations to perform
numsets = numrecs - 11

'step through each regression and nac calculation
For j = 1 To numsets

   'fill n, x(), y() with 12 data points
   n = 0
   meanactpre = 0
   For i = j To j + 11
      If d(i, xfld) <> -99 And d(i, yfld) <> -99 Then
         If numxvars = 1 Or (numxvars = 2 And d(i, 9) <> -99) Or (numxvars = 3 And d(i, 9) <> -99 And d(i, 10) <> -99) Then
                     
            'fill n, x() and y()
            n = n + 1
            x(n, 1) = 1
            x(n, 2) = d(i, xfld)
            If numxvars = 2 Then
               x(n, 3) = d(i, 9)
            ElseIf numxvars = 3 Then
               x(n, 3) = d(i, 9)
               x(n, 4) = d(i, 10)
            End If
            y(n, 1) = d(i, yfld)
            ytot = ytot + y(n, 1)
            meanactpre = ytot / n
            
            'finds max and mins
            If n = 1 Then
               xmin = x(1, 2)
               xmax = x(1, 2)
               ymin = y(1, 1)
               ymax = y(1, 1)
            End If
            If x(n, 2) < xmin Then xmin = x(n, 2)
            If x(n, 2) > xmax Then xmax = x(n, 2)
            If y(n, 1) < ymin Then ymin = y(n, 1)
            If y(n, 1) > ymax Then ymax = y(n, 1)
         End If
      End If
   Next i
   
   'init reg coefs and stats
   xcp1 = -99
   xcp2 = -99
   ycp = -99
   ls = -99
   rs = -99
   ivcoefs(2) = -99
   ivcoefs(3) = -99
   r2 = -99
   rmse = -99
   cvrmse = -99
      
   'calc best fit regression model
   If n > 7 Then
      If nmt = "2P" Then
         'calc model
         Call TwoPMVR(x(), y(), n, numxvars, aa, seaa, bb, sebb, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)
         modeleqn = Format(aa, "#,##0.00") + "  +  " + Format(bb, "#,##0.00") + " T"
         'transfer coefs to std coefs
         xcp1 = -99
         xcp2 = -99
         ycp = aa
         ls = bb
         rs = -99
      ElseIf nmt = "3PC" Or nmt = "3PH" Then
         If nmt = "3PC" Then
            Index = 0 'for 3PC
         ElseIf nmt = "3PH" Then
            Index = 1 'for 3PH
         Else
            Stop
         End If
         Call ThreePMVR(x(), y(), n, numxvars, Index, xcp, sexcp, ycp, seycp, slope, seslope, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)
         'transfers reg coefs to std reg coefs
         If Index = 0 Then '3pc
            ls = 0: sels = 0
            rs = slope: sers = seslope
         Else '3ph
            ls = slope: sels = seslope
            rs = 0: sers = 0
         End If
         modeleqn = Format(ycp, "#,##0.00") + "  -  " + Format(ls, "#,##0.00") + " (" + Format(xcp, "#,##0.00") + " - " + "T" + ")+  +  " + Format(rs, "#,##0.00") + " (" + "T" + " - " + Format(xcp, "#,##0.00") + ")+"
         xcp1 = xcp
         xcp2 = -99
      ElseIf nmt = "4P" Then
         Call FourPMVR(x(), y(), n, numxvars, xcp, sexcp, ycp, seycp, ls, sels, rs, sers, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)
         modeleqn = Format(ycp, "#,##0.00") + "  -  " + Format(ls, "#,##0.00") + " (" + Format(xcp, "#,##0.00") + " - T)+  +  " + Format(rs, "#,##0.00") + " (T - " + Format(xcp, "#,##0.00") + ")+"
         xcp1 = xcp
         xcp2 = -99
      ElseIf nmt = "5P" Then
         Call FivePMVR(x(), y(), n, numxvars, cp1, sexcp1, cp2, sexcp2, ycp, seycp, ls, sels, rs, sers, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)
         modeleqn = Format(ycp, "#,##0.00") + "  -  " + Format(ls, "#,##0.00") + " (" + Format(cp1, "#,##0.00") + " - T)+  +  " + Format(rs, "#,##0.00") + " (T - " + Format(cp2, "#,##0.00") + ")+"
         xcp1 = cp1
         xcp2 = cp2
      ElseIf nmt = "AS" Then
         Stop
      End If
      
   End If
   
   'save reg params as global variables
   nxcp1 = xcp1
   nxcp2 = xcp2
   nycp = ycp
   nls = ls
   nrs = rs
   nivcoefs(2) = ivcoefs(2)
   nivcoefs(3) = ivcoefs(3)
   
   'initialize counters
   numdayspost = 0
   numpostobs = 0
   
   'drive regression model with 12 months of toanorm() to calc nac
   For i = j To j + 11
      startdate = CDate(Format(e(i - 1, 1)) + "/" + Format(e(i - 1, 2)) + "/" + Format(e(i - 1, 3)))
      enddate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
      If d(i, xfld) <> -99 And d(i, yfld) <> -99 And n > 7 Then
         If numxvars = 1 Or (numxvars = 2 And d(i, 9) <> -99) Or (numxvars = 3 And d(i, 9) <> -99 And d(i, 10) <> -99) Then
            'enddate = CDate(Format(e(i, 1)) + "/" + Format(e(i, 2)) + "/" + Format(e(i, 3)))
            
            If nmt = "2P" Then
                       
               'calc predicted energy
               If numxvars = 1 Then
                  d(i, yfldproj) = ycp + ls * toanorm(i) ' + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
               ElseIf numxvars = 2 Then
                  'd(i, yfldproj) = ycp + ls * toanorm(i) + ivcoefs(2) * d(i, 9)     '+ ivcoefs(3) * d(i, 10)
                  d(i, yfldproj) = ycp + ls * toanorm(i) + ivcoefs(2) * tiv1norm(i)     '+ ivcoefs(3) * d(i, 10)
               ElseIf numxvars = 3 Then
                  'd(i, yfldproj) = ycp + ls * toanorm(i) + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
                  d(i, yfldproj) = ycp + ls * toanorm(i) + ivcoefs(2) * tiv1norm(i) + ivcoefs(3) * tiv2norm(i)
               End If
            
            ElseIf nmt = "3PC" Or nmt = "3PH" Then
                    
               'calc predicted energy
               If toanorm(i) <= xcp Then
                 indvar1 = 1
                 indvar2 = 0
               Else
                  indvar1 = 0
                  indvar2 = 1
               End If
               If numxvars = 1 Then
                  d(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) ' + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
               ElseIf numxvars = 2 Then
                  'd(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * d(i, 9)  '+ ivcoefs(3) * d(i, 10)
                  d(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * tiv1norm(i)  '+ ivcoefs(3) * d(i, 10)
               ElseIf numxvars = 3 Then
                  'd(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
                  d(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * tiv1norm(i) + ivcoefs(3) * tiv2norm(i)
               End If
            
            
            ElseIf nmt = "4P" Then
                 
               'calc predicted energy
               If toanorm(i) <= xcp Then
                 indvar1 = 1
                 indvar2 = 0
               Else
                  indvar1 = 0
                  indvar2 = 1
               End If
               If numxvars = 1 Then
                  d(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) ' + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
               ElseIf numxvars = 2 Then
                  'd(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * d(i, 9)  '+ ivcoefs(3) * d(i, 10)
                  d(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * tiv1norm(i)  '+ ivcoefs(3) * d(i, 10)
               ElseIf numxvars = 3 Then
                  'd(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
                  d(i, yfldproj) = ycp - ls * indvar1 * (xcp - toanorm(i)) + rs * indvar2 * (toanorm(i) - xcp) + ivcoefs(2) * tiv1norm(i) + ivcoefs(3) * tiv2norm(i)
               End If
            
            ElseIf nmt = "5P" Then
                              
               'calc predicted energy
               If toanorm(i) <= cp1 Then
                 indvar1 = 1
               Else
                  indvar1 = 0
               End If
               If toanorm(i) <= cp2 Then
                  indvar2 = 0
               Else
                  indvar2 = 1
               End If
                           
               If numxvars = 1 Then
                  d(i, yfldproj) = ycp + ls * indvar1 * (toanorm(i) - cp1) + rs * indvar2 * (toanorm(i) - cp2) ' + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
               ElseIf numxvars = 2 Then
                  'd(i, yfldproj) = ycp + ls * indvar1 * (toanorm(i) - cp1) + rs * indvar2 * (toanorm(i) - cp2) + ivcoefs(2) * d(i, 9)  '+ ivcoefs(3) * d(i, 10)
                  d(i, yfldproj) = ycp + ls * indvar1 * (toanorm(i) - cp1) + rs * indvar2 * (toanorm(i) - cp2) + ivcoefs(2) * tiv1norm(i)  '+ ivcoefs(3) * d(i, 10)
               ElseIf numxvars = 3 Then
                  'd(i, yfldproj) = ycp + ls * indvar1 * (toanorm(i) - cp1) + rs * indvar2 * (toanorm(i) - cp2) + ivcoefs(2) * d(i, 9) + ivcoefs(3) * d(i, 10)
                  d(i, yfldproj) = ycp + ls * indvar1 * (toanorm(i) - cp1) + rs * indvar2 * (toanorm(i) - cp2) + ivcoefs(2) * tiv1norm(i) + ivcoefs(3) * tiv2norm(i)
               End If
            
            ElseIf nmt = "AS" Then
               
               Stop
            
            End If
            
            'sum results
            numdayspost = numdayspost + (enddate - startdate)
            numpostobs = numpostobs + 1
            ac(j + 11) = ac(j + 11) + d(i, yfld)
            nac(j + 11) = nac(j + 11) + d(i, yfldproj)
         
         Else
            d(i, yfldproj) = ndflag
         End If
      Else
         d(i, yfldproj) = ndflag
      End If
      startdate = enddate
   Next i

   'adjusts totsav for other than 365.25 days
   If numdayspost > 0 Then
      nac(j + 11) = nac(j + 11) * 365.25 / numdayspost
      ac(j + 11) = ac(j + 11) * 365.25 / numdayspost
   Else
      nac(j + 11) = 0
      ac(j + 11) = 0
   End If
   
   'save ac, nac and reg coefs in d()
   d(j + 11, 15) = ac(j + 11)
   d(j + 11, 16) = nac(j + 11)
   d(j + 11, 17) = xcp1
   d(j + 11, 18) = xcp2
   d(j + 11, 19) = ycp
   d(j + 11, 20) = ls
   d(j + 11, 21) = rs
   d(j + 11, 22) = ivcoefs(2)
   d(j + 11, 23) = ivcoefs(3)
   
   'calc uncertainty
   ''If nn > 0 And nac > 0 Then
    '  uncert = (1.96 * nrmse * ((1 + 2 / nn) * numpostobs) ^ 0.5) * 365.25 / numdayspost
    '  reluncert = Abs(uncert / nac) * 100
   'Else
   '   uncert = 0
   '   reluncert = 0
   'End If
   
   'graph projection
   ETMain.Graphbox.Cls
   Dim grid, datapnts
   color1 = BLACK
   color2 = BLUE
   grid = True
   datapnts = True
   Call TS2Graph(yfld, yfldproj, grid, datapnts)

   'print baseline model results to status box
   ETMain.StatusBox.Cls
   ETMain.StatusBox.Print "Baseline model stats: "; nmt; "   N = "; nn; "   R2 = "; Format(nr2, "0.00"); "   CV-RMSE = "; Format(ncvrmse, "0.0"); "%"
   If numxvars = 1 Then
      ETMain.StatusBox.Print "Baseline model: "; etype; " = "; modeleqn
   ElseIf numxvars = 2 Then
      ETMain.StatusBox.Print "Baseline model: "; etype; " = "; modeleqn; "  +  "; Format(ivcoefs(2), "#,##0.00"); " IV1"
   ElseIf numxvars = 3 Then
      ETMain.StatusBox.Print "Baseline model: "; etype; " = "; modeleqn; "  +  "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
   End If
   'print ac and nac to status box
   ETMain.StatusBox.Print "Annual Consumption (AC) during baseline period, with actual weather = "; Format(ac(j + 11), "#,##0.0"); " units/year"
   If ac(j + 11) <> 0 Then
      ETMain.StatusBox.Print "Normal Annual Consumption (NAC) during baseline period, if period had normal (TMY2) weather = "; Format(nac(j + 11), "#,##0.0"); " +- "; Format(uncert, "#,##0.0"); " ("; Format(reluncert, "#,##0.0"); "%)"; " units/year     % Change [(NAC-AC)/AC] = "; Format((nac(j + 11) - ac(j + 11)) / ac(j + 11) * 100, "#,##0.0"); "%"
   Else
      ETMain.StatusBox.Print "Normal Annual Consumption (NAC) during baseline period, if period had normal (TMY2) weather = "; Format(nac(j + 11), "#,##0.0"); " +- "; Format(uncert, "#,##0.0"); " ("; Format(reluncert, "#,##0.0"); "%)"; " units/year" '     % Change [(NAC-AC)/AC] = "; Format((nac(j + 11) - ac(j + 11)) / ac(j + 11) * 100, "#,##0.0"); "%"
   End If
   'ETMain.StatusBox.Print "  Annual Consumption = "; Format(ac(j), "#,##0.000"); "  Normalized Annual Consumption = "; Format(nac(j), "#,##0.000"); " +- "; Format(uncert, "#,##0.0"); " ("; Format(reluncert, "#,##0.0"); "%)"
   'ETMain.StatusBox.Print "  Number observations in NAC period = "; Format(numpostobs, "#,##0"); "  Number days in NAC period = "; Format(numdayspost, "#,##0")

   'print sitename and pre model coefs output file
   'If msslidinganal <> True Then 'single-site
      Print #9, efilename$ + Chr(9); 'sitename m(i,1)
   'Else 'multi-site
   '   Print #9, m(i, 1) + Chr(9); 'sitename m(i,1)
   'End If
   Print #9, nmt + Chr(9); 'modeltype
   Print #9, Format$(enddate) + Chr(9); 'end date of 12 month period
   Print #9, Format$(n, "#,##0") + Chr(9);
   Print #9, Format$(meanactpre, "#,##0.00") + Chr(9);
   Print #9, Format$(r2, "0.0000") + Chr(9);
   Print #9, Format$(rmse, "#,##0.0000") + Chr(9);
   Print #9, Format$(cvrmse, "#,##0.0000") + Chr(9);
   Print #9, Format$(xcp1, "#,##0.0000") + Chr(9);
   Print #9, Format$(xcp2, "#,##0.0000") + Chr(9);
   Print #9, Format$(ycp, "#,##0.0000") + Chr(9);
   Print #9, Format$(ls, "#,##0.0000") + Chr(9);
   Print #9, Format$(rs, "#,##0.0000") + Chr(9);
   Print #9, Format$(ivcoefs(2), "#,##0.0000") + Chr(9);
   Print #9, Format$(ivcoefs(3), "#,##0.0000") + Chr(9);

   'print savings summary to output file
   If acg = Empty Then acg = -99
   If nacg = Empty Then nacg = -99
   Print #9, Format$(ac(j + 11), "#,##0.00") + Chr(9);
   Print #9, Format$(nac(j + 11), "#,##0.00") + Chr(9)
   'Print #9, Format$(numpostobs, "#,##0") + Chr(9);
   'Print #9, Format$(meanactpost, "#,##0.00") + Chr(9);
   'Print #9, Format$(numdayspost, "#,##0") + Chr(9);
   'Print #9, Format$(totsav, "#,##0") + Chr(9);
   'Print #9, Format$(uncert, "#,##0.0") + Chr(9);
   'Print #9, Format$(percentsav, "0") + Chr(9);
   'Print #9, Format(reluncert, "#,##0.0") + Chr(9);
   'Print #9, Format$(savperobs, "#,##0.00") + Chr(9);
   'Print #9, Format(savperday, "#,##0.000")
  
   'save results from first set
   'If j = 1 Then
   '   acfirst = ac(j + 11)
   '   nacfirst = nac(j + 11)
   'End If
   'aclast = ac(j + 11)
   'naclast = nac(j + 11)
      
   'calc total nac to use in calc of avg nac
   If nac(j + 11) > 0 And nac(j + 11) <> -99 Then
      numnacs = numnacs + 1
      NACTot = NACTot + nac(j + 11)
      If Index = 0 Then '3pc
         STot = STot + d(12, 21) 'rs
      Else '3ph
         STot = STot + d(12, 20) 'ls
      End If
      TbalTot = TbalTot + d(j + 11, 17)
      EindTot = EindTot + d(j + 11, 19)
   End If

Next j

'save ac, nac and reg coefs in d()
'd(j + 11, 15) = ac(j + 11)
'd(j + 11, 16) = nac(j + 11)
'd(j + 11, 17) = xcp1
'd(j + 11, 18) = xcp2
'd(j + 11, 19) = ycp
'd(j + 11, 20) = ls
'd(j + 11, 21) = rs
'd(j + 11, 22) = ivcoefs(2)
'd(j + 11, 23) = ivcoefs(3)

ACfirst = d(12, 15)
NACfirst = d(12, 16)
Tbalfirst = d(12, 17)
Eindfirst = d(12, 19)
If Index = 0 Then '3pc
   Sfirst = d(12, 21) 'rs
Else '3ph
   Sfirst = d(12, 20) 'ls
End If

NAClast_1 = d(j + 9, 16)
Tballast_1 = d(j + 9, 17)
Eindlast_1 = d(j + 9, 19)
If Index = 0 Then '3pc
   Slast_1 = d(j + 9, 21) 'rs
Else '3ph
   Slast_1 = d(j + 9, 20) 'ls
End If

AClast = d(j + 10, 15)
NAClast = d(j + 10, 16)
Tballast = d(j + 10, 17)
Eindlast = d(j + 10, 19)
If Index = 0 Then '3pc
   Slast = d(j + 10, 21) 'rs
Else '3ph
   Slast = d(j + 10, 20) 'ls
End If

'calc avg nac
If numnacs > 0 Then
   NACmean = NACTot / numnacs
   Tbalmean = TbalTot / numnacs
   Eindmean = EindTot / numnacs
   Smean = STot / numnacs
Else
   NACmean = -99
   Tbalmean = -99
   Eindmean = -99
   Smean = -99
End If

'graph results
vn$(15) = "Annual Consumption"
vn$(16) = "Normal Annual Consumption"
grpfld = 0
'graph ac and nacs
ETMain.Graphbox.Cls
'Dim grid, datapnts
color1 = BLACK
color2 = BLUE
grid = True
datapnts = True
Call TS2Graph(16, 15, grid, datapnts)

'print summary results to screen
ETMain.StatusBox.Cls
ETMain.StatusBox.Print "Energy filename: "; efilename$
ETMain.StatusBox.Print "Sliding baseline model: "; nmt '; "   N = "; nn; "   R2 = "; Format(nr2, "0.00"); "   CV-RMSE = "; Format(ncvrmse, "0.0"); "%"
If ACfirst > 0 And NACfirst > 0 Then
   ETMain.StatusBox.Print "First and last Annual Consumption (AC) during baseline period, with actual weather = "; Format(ACfirst, "#,##0.0"); " units/year and "; Format(AClast, "#,##0.0"); " units/year"; "     % Change [(AC1-AC2)/AC1] = "; Format((ACfirst - AClast) / ACfirst * 100, "#,##0.0"); "%"
   ETMain.StatusBox.Print "First and last Normal Annual Consumption (NAC) during baseline period, if period had normal (TMY2) weather = "; Format(NACfirst, "#,##0.0"); " units/year and "; Format(NAClast, "#,##0.0"); " units/year"; "     % Change [(NAC1-NAC2)/NAC1] = "; Format((NACfirst - NAClast) / NACfirst * 100, "#,##0.0"); "%"
Else
   ETMain.StatusBox.Print "First and last Annual Consumption (AC) during baseline period, with actual weather = "; Format(ACfirst, "#,##0.0"); " units/year and "; Format(AClast, "#,##0.0"); " units/year" '; "     % Change [(AC1-AC2)/AC1] = "; Format((acfirst - aclast) / acfirst * 100, "#,##0.0"); "%"
   ETMain.StatusBox.Print "First and last Normal Annual Consumption (NAC) during baseline period, if period had normal (TMY2) weather = "; Format(NACfirst, "#,##0.0"); " units/year and "; Format(NAClast, "#,##0.0"); " units/year" '; "     % Change [(NAC1-NAC2)/NAC1] = "; Format((nacfirst - naclast) / nacfirst * 100, "#,##0.0"); "%"
End If

csce:
Screen.MousePointer = 0
Unload PerDone
If msslidinganal <> True Then
   Close #9
End If
'prints error message
If Err Then
   msg$ = """" + Error(Err) + """"
   MsgBox msg$, , "Error"
   Resume cses
End If
cses:

End Sub



Public Sub FillArray(filename$, arr(), anumrecs, anumflds)
Dim l$, delim, charnum
Dim i, j, scharnum, fnum, echarnum, flength, row, col
'Dim anumrecs, anumflds

'sets default no data flag
ndflag = -99

'finds delimiter type, anumflds, anumrecs
anumrecs = 0
anumflds = 0
Open filename$ For Input As #1
While Not EOF(1)
   'reads a line
   Line Input #1, l$
   l$ = Trim$(l$)
   If (Len(l$) > 0) And (Left(l$, 1) = "1" Or Left(l$, 1) = "2" Or Left(l$, 1) = "3" Or Left(l$, 1) = "4" Or Left(l$, 1) = "5" Or Left(l$, 1) = "6" Or Left(l$, 1) = "7" Or Left(l$, 1) = "8" Or Left(l$, 1) = "9" Or Left(l$, 1) = "0") Then
      anumrecs = anumrecs + 1
   
      If anumrecs = 1 Then
         
         'determines if comma, tab (chr(9)), or space (chr(32)) delimited
         delim = ""
         For charnum = 1 To Len(l$)
            If Mid$(l$, charnum, 1) = "," Then
               delim = ","
               Exit For
            End If
         Next charnum
         If delim <> "," Then
            For charnum = 1 To Len(l$)
               If Mid$(l$, charnum, 1) = Chr(9) Then
                  delim = Chr(9)
                  Exit For
               End If
            Next charnum
         End If
         If delim = "" Then delim = Chr(32)
   
         'determines anumflds and dims field()
         anumflds = 0
         If delim = "," Or delim = Chr(9) Then
            For charnum = 1 To Len(l$)
               If Mid$(l$, charnum, 1) = delim Or charnum = Len(l$) Then
                  anumflds = anumflds + 1
               End If
            Next charnum
         Else 'delim = " "
            For charnum = 1 To Len(l$)
               If Mid$(l$, charnum, 1) = delim Then
                  If Mid$(l$, charnum - 1, 1) <> delim Then anumflds = anumflds + 1
               ElseIf charnum = Len(l$) Then
                  anumflds = anumflds + 1
               End If
            Next charnum
         End If
      End If
   End If
Wend
Close #1
ReDim arr(anumrecs, anumflds)


'fills arr(row,col) if space, tab or comma delim
row = 0
Open filename$ For Input As #1
While Not EOF(1)
   'reads a line
   Line Input #1, l$
   l$ = Trim$(l$)
   If (Len(l$) > 0) And (Left(l$, 1) = "1" Or Left(l$, 1) = "2" Or Left(l$, 1) = "3" Or Left(l$, 1) = "4" Or Left(l$, 1) = "5" Or Left(l$, 1) = "6" Or Left(l$, 1) = "7" Or Left(l$, 1) = "8" Or Left(l$, 1) = "9" Or Left(l$, 1) = "0") Then
      row = row + 1
      'parses the line into fields
      scharnum = 1
      fnum = 0
      For charnum = 1 To Len(l$)
         If Mid$(l$, charnum, 1) = delim Then
            If Mid$(l$, charnum - 1, 1) <> delim Then
               fnum = fnum + 1
               If fnum > anumflds Then
                  msg$ = "Incorrect number of columns in line " + Str$(row) + " of " + UCase$(filename$) + "."
                  MsgBox msg$, , "Error"
                  Close #1
                  Screen.MousePointer = 0
                  Exit Sub
               End If
               echarnum = charnum - 1
               flength = echarnum - scharnum + 1
               arr(row, fnum) = Val(Trim(Mid$(l$, scharnum, flength)))
               scharnum = echarnum + 2
            End If
         ElseIf charnum = Len(l$) Then
            fnum = fnum + 1
            echarnum = charnum
            flength = echarnum - scharnum + 1
            arr(row, fnum) = Val(Trim(Mid$(l$, scharnum, flength)))
         End If
      Next charnum
   End If
   Call UpdatePerDone(row, anumrecs * 2, 1)
Wend

Close #1
End Sub
