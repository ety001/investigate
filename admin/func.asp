<%
Sub Debug_noEnd(a,b)
	response.Write(a & ":" & b)
	response.Write("<br />")
End Sub
Sub Debug_End(a,b)
	response.Write(a & ":" & b)
	response.End()
End Sub

%>
