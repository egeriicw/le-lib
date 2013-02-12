Attribute VB_Name = "ETGlobal"
Option Explicit
'*******************For ETSubs*********************************
'fileopen
Global tnp$, tfn$, opencode

'open energy
Global numindvars

'data
Global e(), w(), erecs, eflds, wrecs, wflds, eopen, wopen, tivopen

'for editor
Global filesaved, fscancel$, newfile, formcap
Global opath$, ofilename$, focancel$
Global findcancel, findtxt$, txtchanged, lastposition, txtposition
Global nsno, nscancel, nsyes, msg$

'to print results to output file
Global efilepath$, efilename$

'for multisite analysis
Global mfilepath$, mfilename$, mrecs, mflds, m(), skip
Global nnumdayspost, nnumpostobs, ntotsav, nuncert, npercentsav, nreluncert, nsavperobs, nsavperday
Global nmeanactpre, nmeanactpost
Global engytype$, modeltype$, mslopen, msslidinganal

'for modeling
Global nmt, nrmse, nn, ncvrmse, nr2
Global nxcp1, nxcp2, nycp, nls, nrs, nivcoefs(), modeleqn

'for nac
Global tm2open, nacg, acg, uncertg, reluncertg
Global nac(), ac(), nxcp1s(), nxcp2s(), nycps(), nlss(), nrss(), nivcoef1s(), nivcoef2s()
Global ACfirst, AClast
Global NACfirst, NAClast, NAClast_1, NACmean
Global Sfirst, Slast, Slast_1, Smean
Global Tbalfirst, Tballast, Tballast_1, Tbalmean
Global Eindfirst, Eindlast, Eindlast_1, Eindmean
Global tiv()
Global tivrecs, tivflds


'******************For GSubs*************************************
'opendata
'input: make necessary changes in opendata
'output: sets or fills these global variables
Global d(), numrecs, numflds, vn$(), ndflag
Global mofld, dyfld, yrfld, hrfld, timeint

'setall
'input:  everything from open data
'output: sets or fills these global variables
Global n, x(), y(), grpxy()
Global ymin, ymax, yrange, yminr, ymaxr, yranger, yroot, ypower, yfmt, yintindex
Global xmin, xmax, xrange, xminr, xmaxr, xranger, xroot, xpower, xfmt, xintindex

'coef matrix for drawing model lines on xy plots
Global c() 'c(3 energy types, 5 model coefficents)

'optional graphing variables for grouping
'grpfld is a fld in d() which contains a 1 or 2
'(if grpfld is defined then) if d(grpfld) = 1 then color = color1 else color = color2
Global grpfld, color1, color2

'TS3DGraph
'these are needed for TS3DGraph to communicate with it's client subroutines
Global res, maskmin(), maskmax(), maskminnew(), maskmaxnew()
Global xclick, yclick, sint, cost, numdays, dy1, thetad

'if you want click-identification, then these must be global and set
Global gtype$, xvar, yvar1, yvar2, xcur, ycur

' BackColor, ForeColor, FillColor (standard RGB colors: form, controls)
Global Const BLACK = &H0&
Global Const RED = &HFF&
Global Const GREEN = &HFF00&
Global Const YELLOW = &HFFFF&
Global Const BLUE = &HFF0000
Global Const MAGENTA = &HFF00FF
Global Const CYAN = &HFFFF00
Global Const WHITE = &HFFFFFF





