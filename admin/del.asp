<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
If Session("Logined") <> "yes" Then
	response.redirect("login.asp")
Else
%>
<!--#include file="func.asp"-->
<!--#include file="../Include/conn.asp"-->
<%
title = request.QueryString("title")
QN_ID = request.QueryString("QN_ID")

DelSql = "delete * from QN_Title where Title='" & title & "'"
conn.execute(DelSql)

DelSql = "delete * from YON where YON='" & QN_ID & "'"
conn.execute(DelSql)

DelSql = "drop table Q" & QN_ID
conn.execute(DelSql)

DelSql = "drop table A" & QN_ID
conn.execute(DelSql)

conn.close
Set conn = nothing

response.Redirect("viewList.asp")

End If
%>