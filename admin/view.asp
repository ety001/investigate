<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
If Session("Logined") <> "yes" Then	'检查登陆情况
	response.redirect("login.asp")
Else
%>
<!--#include file="../Include/conn.asp"-->
<%
title = request.QueryString("title")
Set rsList = Server.CreateObject("Adodb.RecordSet")
ListSql = "select * from QN_Title " & "where title='" & title & "'" 
rsList.Open ListSql,conn,1,1
QN_ID = rsList("QN_ID")
rsList.close
Set rsList = nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<title>当前显示的是 <%=title%> 问卷的内容</title>
</head>
<%
Set rs = Server.CreateObject("Adodb.RecordSet")
SqlStr = "select * from Q" & QN_ID & " order by QNum asc"
rs.Open SqlStr,conn,1,1
%>
<body>
<div class="top" style="margin:0 auto; width=960px;">当前显示的是 <b><%=title%></b> 问卷的内容&nbsp;<a href="viewList.asp">返回上一级</a></div>
<div class="QTable" style="margin:0 auto; width=960px;">

		<%Do While Not rs.eof%>
    	<%
			QType = rs("QType")
			Option_Num = rs("Option_num")
		%>
        <%Select Case Qtype%>
			<%Case "single"%>
            	<div class="question">
                <div class="Qtitle"><%=rs("QNum")%>.<%=rs("QTitle")%></div>
        		<div class="Qanswer">
            	<%For i = 1 to Option_Num %>
            		<label><input name="Q<%=rs("QNum")%>" type="radio" value="<%=chr(64+i)%>" /><%=chr(64+i)%>.<%=rs("Op" & i)%></label>
            	<%Next%>
            	</div>
                </div>
                
            <%Case "multiple"%>
       			<div class="question">
                <div class="Qtitle"><%=rs("QNum")%>.<%=rs("QTitle")%></div>
            	<div class="Qanswer">
                <%For i = 1 to Option_Num %>
                	<label><input name="Q<%=rs("QNum")%>" type="checkbox" value="<%=chr(64+i)%>" /><%=chr(64+i)%>.<%=rs("Op" & i)%></label>
                <%Next%>
                </div>
                </div>
            <%Case "question"%>
            	<div class="question">
                <div class="Qtitle"><%=rs("QNum")%>.<%=rs("Question")%></div>
                <div class="Qanswer"><label><textarea name="Q<%=rs("QNum")%>" cols="50" rows="10"></textarea></label></div>
                </div>
        <%End Select%>
        <%
		rs.movenext
		Loop
		rs.close
		Set rs = nothing
		conn.close
		Set conn = nothing
		%>

</div>
</body>
</html>
<%End If'判断登陆条件的结束标记%>