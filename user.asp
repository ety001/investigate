<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
If request.QueryString("q") = "quit" Then
	Session("User_Logined") = "no"
	Session("Stu_ID") = ""
	response.Redirect("login.asp")
End If
If Session("Stu_ID") = "" and Session("User_Logined") <> "yes" Then
	response.redirect("login.asp")
Else
%>
<!--#include file="Include/user_conn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<title><%=Session("Stu_name")%> ͬѧ�Ĳ������桪��³����ѧ�����ʾ�ϵͳ Power By ETY001</title>
<style type="text/css">
<!--
td {
	text-align: center;
}
-->
</style>
</head>
<body><center>
<div style="width:724px; margin:0 auto; text-align:right;"><a href="user.asp?q=quit">�˳�</a></div>
<div style="padding:5px 10px; width:724px; margin:0 auto; font-family:'����'; font-size:20px; font-weight:bold; color:#F00;">ע�⣺ÿ�ݵ����ʾ�ֻ����дһ�Σ��������д��</div>
<table width="729" border="0" bgcolor="#666666">
  <tr>
    <th width="146" scope="col" bgcolor="#CCFF33">�����ʾ���</th>
    <th width="146" scope="col" bgcolor="#CCFF33">�����ʾ�����</th>
    <th width="415" scope="col" bgcolor="#CCFF33">�����ʾ����</th>
  </tr>
<%
Set rsList = Server.CreateObject("Adodb.RecordSet")
ListSql = "select * from QN_Title"
rsList.Open ListSql,conn,1,1
i = 1
Do While Not rsList.Eof
	If Not rsList.Eof Then
%>
  <tr>
  <%If i mod 2 = 0 Then%>
  	<td bgcolor="#CCCCCC"><%=rsList("QN_ID")%></td>
    <td bgcolor="#CCCCCC"><%=rsList("QNum")%></td>
    <td bgcolor="#CCCCCC"><b><a href="index.asp?title=<%=rsList("Title")%>&QN_ID=<%=rsList("QN_ID")%>" target="_blank"><%=rsList("Title")%></a></b></td>
   <%Else%>
   	<td bgcolor="#66CCCC"><%=rsList("QN_ID")%></td>
    <td bgcolor="#66CCCC"><%=rsList("QNum")%></td>
    <td bgcolor="#66CCCC"><b><a href="index.asp?title=<%=rsList("Title")%>&QN_ID=<%=rsList("QN_ID")%>" target="_blank"><%=rsList("Title")%></a></b></td>
   <%End If%>
  </tr>
<%
		rsList.movenext
		i = i + 1
	End If		
Loop
%>
</table>
</center>
</body>
</html>
<%End If%>