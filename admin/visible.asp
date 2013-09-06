<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
If Session("Logined") <> "yes" Then
	response.redirect("login.asp")
Else
%>
<!--#include file="func.asp"-->
<!--#include file="../Include/conn.asp"-->
<%
key = request.QueryString("key")
QN_ID = request.QueryString("QN_ID")
If key = "yes" Then
	ChgSql = "update QN_Title set QVisible='1' where QN_ID='" & QN_ID & "'"
End If
If key = "no" Then
	ChgSql = "update QN_Title set QVisible='0' where QN_ID='" & QN_ID & "'"
End If
conn.execute(ChgSql)

response.Redirect("viewList.asp")

End If
%>