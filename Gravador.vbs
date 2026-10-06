' Gravador SAP GUI: grava as ações feitas na sessão ativa e gera VBScript + Python.
' Uso: dê dois cliques, faça as ações no SAP e clique OK na caixa "Gravando..." para parar.
Option Explicit

Dim sapGui, app, session, vbsLines, pyLines, nCmds, fso, baseDir, stamp

On Error Resume Next
Set sapGui = GetObject("SAPGUI")
If Err.Number <> 0 Then
  MsgBox "SAP GUI não encontrado ou scripting desabilitado.", vbExclamation, "Gravador SAP"
  WScript.Quit 1
End If
Set app = sapGui.GetScriptingEngine
Set session = app.ActiveSession
If session Is Nothing Or Err.Number <> 0 Then
  Err.Clear
  Set session = app.Children(0).Children(0)
End If
If Err.Number <> 0 Then
  MsgBox "Nenhuma sessão SAP aberta.", vbExclamation, "Gravador SAP"
  WScript.Quit 1
End If
On Error GoTo 0

vbsLines = "" : pyLines = "" : nCmds = 0

WScript.ConnectObject session, "on_"
session.Record = True

MsgBox "Gravando a sessão " & session.Id & " (" & session.Info.Transaction & ")." & vbCrLf & vbCrLf & _
       "Faça as ações no SAP. Clique OK aqui para PARAR.", vbInformation + vbSystemModal, "Gravando..."

session.Record = False
WScript.DisconnectObject session

If nCmds = 0 Then
  MsgBox "Nenhuma ação gravada.", vbInformation, "Gravador SAP"
  WScript.Quit 0
End If

Set fso = CreateObject("Scripting.FileSystemObject")
baseDir = fso.GetParentFolderName(WScript.ScriptFullName) & "\gravacoes"
If Not fso.FolderExists(baseDir) Then fso.CreateFolder baseDir
stamp = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & _
        Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)

WriteFile baseDir & "\gravacao_" & stamp & ".vbs", _
  "Set SapGui = GetObject(""SAPGUI"")" & vbCrLf & _
  "Set app = SapGui.GetScriptingEngine" & vbCrLf & _
  "Set session = app.ActiveSession" & vbCrLf & vbCrLf & vbsLines & vbCrLf

WriteFile baseDir & "\gravacao_" & stamp & ".py", _
  "import win32com.client" & vbCrLf & vbCrLf & _
  "sap = win32com.client.GetObject(""SAPGUI"")" & vbCrLf & _
  "app = sap.GetScriptingEngine" & vbCrLf & _
  "session = app.ActiveSession" & vbCrLf & vbCrLf & pyLines & vbCrLf

MsgBox nCmds & " comandos gravados em:" & vbCrLf & baseDir, vbInformation, "Gravador SAP"


' ---- Evento disparado pelo SAP a cada ação do usuário (session.Record = True) ----
Sub on_Change(ses, comp, cmdArray)
  Dim i, c, rel, kind, name, j, argsV, argsP
  On Error Resume Next
  rel = RelId(comp.Id)
  For i = LBound(cmdArray) To UBound(cmdArray)
    c = cmdArray(i)
    kind = c(0) : name = c(1)
    Select Case kind
      Case "SP"   ' set property
        AddV "session.findById(""" & rel & """)." & name & " = " & FmtV(c(2))
        AddP "session.findById(""" & rel & """)." & name & " = " & FmtP(c(2))
      Case "M"    ' method call
        argsV = "" : argsP = ""
        For j = 2 To UBound(c)
          If j > 2 Then argsV = argsV & ", " : argsP = argsP & ", "
          argsV = argsV & FmtV(c(j)) : argsP = argsP & FmtP(c(j))
        Next
        If argsV = "" Then
          AddV "session.findById(""" & rel & """)." & name
        Else
          AddV "session.findById(""" & rel & """)." & name & " " & argsV
        End If
        AddP "session.findById(""" & rel & """)." & name & "(" & argsP & ")"
      Case Else
        AddV "' [" & kind & "] " & rel & " " & name
        AddP "# [" & kind & "] " & rel & " " & name
    End Select
  Next
End Sub

Function RelId(id)
  Dim p, q
  p = InStr(id, "/ses[")
  If p = 0 Then RelId = id : Exit Function
  q = InStr(p, id, "]/")
  If q = 0 Then RelId = "" Else RelId = Mid(id, q + 2)
End Function

Function FmtV(v)
  Select Case VarType(v)
    Case vbString:  FmtV = """" & Replace(v, """", """""") & """"
    Case vbBoolean: If v Then FmtV = "True" Else FmtV = "False"
    Case Else:      FmtV = CStr(v)
  End Select
End Function

Function FmtP(v)
  Select Case VarType(v)
    Case vbString:  FmtP = """" & Replace(Replace(v, "\", "\\"), """", "\""") & """"
    Case vbBoolean: If v Then FmtP = "True" Else FmtP = "False"
    Case Else:      FmtP = CStr(v)
  End Select
End Function

Sub WriteFile(path, text)
  Dim st
  Set st = CreateObject("ADODB.Stream")
  st.Type = 2 : st.Charset = "utf-8" : st.Open
  st.WriteText text
  st.SaveToFile path, 2
  st.Close
End Sub

Sub AddV(s)
  vbsLines = vbsLines & s & vbCrLf
  nCmds = nCmds + 1
End Sub

Sub AddP(s)
  pyLines = pyLines & s & vbCrLf
End Sub
