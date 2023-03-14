Attribute VB_Name = "Module1"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long



Public mCapHwnd As Long



Public Const CONNECT As Long = 1034

Public Const DISCONNECT As Long = 1035

Public Const GetObject As Long = 1036

Public Const COPY As Long = 1054
Public co As ADODB.Connection
Public cl As ADODB.Recordset
Public mt As ADODB.Recordset
Public em As ADODB.Recordset
Public et As ADODB.Recordset
Public sr As ADODB.Recordset
Public nt As ADODB.Recordset
Public cf As ADODB.Recordset
Public cf1 As ADODB.Recordset
Public ab As ADODB.Recordset
Public pr As ADODB.Recordset
Public ps As ADODB.Recordset
Public cd As ADODB.Recordset
Public ce As ADODB.Recordset
Public rc As ADODB.Recordset
Public ca As ADODB.Recordset
Public jr As ADODB.Recordset
Public pf As ADODB.Recordset
Public dp As ADODB.Recordset
Public pa As ADODB.Recordset
Public pp As ADODB.Recordset
Public bn As ADODB.Recordset
Public cr As ADODB.Recordset
Public an As ADODB.Recordset
Public ut As ADODB.Recordset
Public ev As ADODB.Recordset
Public fc As ADODB.Recordset
Public pfc As ADODB.Recordset
Dim ane As String
Public Enum Ahmede
    arabic = vbMsgBoxRight + vbMsgBoxRtlReading
End Enum
Function cont()
Set co = New ADODB.Connection
Set cl = New ADODB.Recordset
Set mt = New ADODB.Recordset
Set em = New ADODB.Recordset
Set et = New ADODB.Recordset
Set sr = New ADODB.Recordset
Set nt = New ADODB.Recordset
Set cf = New ADODB.Recordset
Set cf1 = New ADODB.Recordset
Set ab = New ADODB.Recordset
Set pr = New ADODB.Recordset
Set ps = New ADODB.Recordset
Set cd = New ADODB.Recordset
Set ce = New ADODB.Recordset
Set rc = New ADODB.Recordset
Set ca = New ADODB.Recordset
Set jr = New ADODB.Recordset
Set pf = New ADODB.Recordset
Set dp = New ADODB.Recordset
Set pa = New ADODB.Recordset
Set pp = New ADODB.Recordset
Set bn = New ADODB.Recordset
Set cr = New ADODB.Recordset
Set an = New ADODB.Recordset
Set ut = New ADODB.Recordset
Set ev = New ADODB.Recordset
Set fc = New ADODB.Recordset
Set pfc = New ADODB.Recordset
If start.Label1.Caption <> "" Then
ane = start.Label1.Caption
'ane = "2012-2013"
Else
ane = face.SBB1.Panels(9).Text
End If
co.Provider = "microsoft.jet.oledb.4.0; jet oledb:database password=7346804"
co.ConnectionString = App.Path & "\" & ane & ".mdb"
co.Open
cl.Open "select*from Tclasses order by cla ASC", co, adOpenKeyset, adLockOptimistic
mt.Open "select*from Tmatiers order by aut ASC", co, adOpenKeyset, adLockOptimistic
em.Open "select*from Templois", co, adOpenKeyset, adLockOptimistic
et.Open "select*from Tetudiants", co, adOpenKeyset, adLockOptimistic
sr.Open "select*from Tseries", co, adOpenKeyset, adLockOptimistic
nt.Open "select*from Tnotes", co, adOpenKeyset, adLockOptimistic
cf.Open "select*from Tcoffdevoirs", co, adOpenKeyset, adLockOptimistic
cf1.Open "select*from Tcoffdevoirs1", co, adOpenKeyset, adLockOptimistic
ab.Open "select*from Tabsences", co, adOpenKeyset, adLockOptimistic
pr.Open "select*from Tprofesseurs", co, adOpenKeyset, adLockOptimistic
ps.Open "select*from Tpresences order by mois ASC", co, adOpenKeyset, adLockOptimistic
cd.Open "select*from Tcodes", co, adOpenKeyset, adLockOptimistic
ce.Open "select*from Tcompteetudiants", co, adOpenKeyset, adLockOptimistic
rc.Open "select*from Trecus order by aut ASC", co, adOpenKeyset, adLockOptimistic
ca.Open "select*from Tcaisse", co, adOpenKeyset, adLockOptimistic
jr.Open "select*from Tjournal", co, adOpenKeyset, adLockOptimistic
pf.Open "select*from Tpayprofesseurs", co, adOpenKeyset, adLockOptimistic
dp.Open "select*from Tdepenses", co, adOpenKeyset, adLockOptimistic
pa.Open "select*from Tpartenaires", co, adOpenKeyset, adLockOptimistic
pp.Open "select*from Tpaypartenaires", co, adOpenKeyset, adLockOptimistic
bn.Open "select*from Tbank", co, adOpenKeyset, adLockOptimistic
cr.Open "select*from Tcontrolerecu", co, adOpenKeyset, adLockOptimistic
an.Open "select*from Tannees", co, adOpenKeyset, adLockOptimistic
ut.Open "select*from Tutilisateurs", co, adOpenKeyset, adLockOptimistic
'ev.Open "select*from Envoyer", co, adOpenKeyset, adLockOptimistic
fc.Open "select*from Tfonctionnaires order by mtr ASC", co, adOpenKeyset, adLockOptimistic
pfc.Open "select*from Tpayfonctionnaires order by aut ASC", co, adOpenKeyset, adLockOptimistic
End Function

