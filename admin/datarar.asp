<!--#include file="adminconn.asp"-->
<!--#include file="admincheck.asp"-->
<%
Response.Expires = -1  
Response.ExpiresAbsolute = Now() - 1  
Response.cachecontrol = "no-cache" 
DIm Objfso
Objfso = "Scripting.FileSystemObject"
Dim Db
Db=admindataurl '这里请根据你实际数据库名称替换
Call Compact

Sub Compact
  Response.Write CompactDB(Server.Mappath(db),false)
End Sub
'=====================压缩数据库=========================
Function CompactDB(dbPath, boolIs97)
 On Error Resume Next
 Dim fso, Engine, strDBPath,JET_3X
 strDBPath = left(dbPath,instrrev(DBPath,"\"))
 Set fso = CreateObject(Objfso)
 If Err Then
    Err.Clear
    CompactDB = Lang.item("g_110") & vbCrLf
    Exit Function
 End If
 If fso.FileExists(dbPath) Then
    fso.CopyFile dbpath,strDBPath & "temp.mdb"
    Set Engine = CreateObject("JRO.JetEngine")

    If boolIs97 = "True" Then
    Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;" _
   & "Jet OLEDB:Engine Type=" & JET_3X
  Else
   Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"
  End If

  fso.CopyFile strDBPath & "temp1.mdb",dbpath
  fso.DeleteFile(strDBPath & "temp.mdb")
  fso.DeleteFile(strDBPath & "temp1.mdb")
  Set fso = Nothing
  Set Engine = Nothing
  '写入日志
dim SQLcmd
SQLcmd = "Insert Into caozuorizhi(cuozuo,caozuoren,cuozuotime,cuozuoip) Values"
SQLcmd = SQLcmd & "('压缩和修复数据库','" & session("admin") & "',now,'" & request.servervariables("REMOTE_HOST") & "')"
conn.Execute SQLcmd
'写入日志结束
  Response.Write "压缩成功!"
 Else
  Response.Write "没搞定,再来一次?"
 End If
End Function
%>
