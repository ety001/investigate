<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
If Session("Logined") <> "yes" Then
	response.redirect("login.asp")
Else
%>
<!--#include file="../Include/conn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<title>查看调查问卷</title>
</head>
<body>
<center>
<a href="index.asp" target="_self">返回</a>
<table width="960" border="1">
  <tr>
    <th width="146" scope="col">调查问卷编号</th>
    <th width="146" scope="col">调查问卷题数</th>
    <th width="415" scope="col">调查问卷标题</th>
    <th width="111" scope="col">是否删除</th>
    <th width="108" scope="col">是否可见</th>
  </tr>
<%
Set rsList = Server.CreateObject("Adodb.RecordSet")
ListSql = "select * from QN_Title"
rsList.Open ListSql,conn,1,1
Do While Not rsList.Eof
	If Not rsList.Eof Then
%>
  <tr>
  	<td><%=rsList("QN_ID")%></td>
    <td><%=rsList("QNum")%></td>
    <td><a href="view.asp?title=<%=rsList("Title")%>" target="_self"><%=rsList("Title")%></a></td>
    <td><a href="del.asp?title=<%=rsList("Title")%>&QN_ID=<%=rsList("QN_ID")%>" target="_self">删除</a></td>
    <td><%If rsList("QVisible") = "1" Then%><b><a href="visible.asp?key=yes&QN_ID=<%=rsList("QN_ID")%>" target="_self">可见</a></b><%else%><a href="visible.asp?key=yes&QN_ID=<%=rsList("QN_ID")%>" target="_self">可见</a></b><%End If%>&amp;<%If rsList("QVisible") = "0" Then%><b><a href="visible.asp?key=no&QN_ID=<%=rsList("QN_ID")%>" target="_self">不可见</a></b><%Else%><a href="visible.asp?key=no&QN_ID=<%=rsList("QN_ID")%>" target="_self">不可见</a><%End IF%></td>
  </tr>
<%
		rsList.movenext
	Else
%>
  <tr>
  	<td><%="暂无信息"%></td>
    <td><%="暂无信息"%></td>
    <td><%="暂无信息"%></td>
    <td><%="暂无信息"%></td>
    <td><%="暂无信息"%></td>
  </tr>
<%
	End If		
Loop
%>
</table>
</center>
</body>
</html>
<%
End If'判断登陆的结束If标记
%>
