<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="func.asp"-->
<%
If Session("Logined") <> "yes" Then	'����½���
	response.redirect("login.asp")
Else
%>
<!--#include file="../Include/conn.asp"-->
<%
QNum = CInt(request.Form("num"))	'�����ʾ���Ŀ����Ŀ
QTitle = request.Form("top")	'�����ʾ����Ŀ
QN_ID = request.Form("QN_ID")	'�����ʾ�ı��
questionnaire = split(trim(request.Form),"&")	'��ȡ������Ϣ���ԡ�&���ŷָ�
Dim subjectTemp(99,10) 	'����һ���ܴ洢100��������Ŀ��ÿ����Ŀ�����9��ѡ������飨��άֵ0������Ŀ�����ͣ�1������Ŀ������������
Dim subject(99,10) 	'����һ���ܴ洢100��������Ŀ��ÿ����Ŀ�����9��ѡ������飨��άֵ0������Ŀ�����ͣ�1������Ŀ������������
subjectNum = 0		'Subject����ĵ�һά����
subjectSelectNum = 2	'Subject����ĵڶ�ά����
upSubject = Ubound(questionnaire)	'���Ա�����Ϣ��Ŀ
'��ѭ���ǳ����ص�
For i = 1 to QNum
	title = "Q" & i
	jj = 2
	For j = 2 to upSubject		'i������¼����form�ύ���ֶ���
		QTemp = Split(unescape(questionnaire(j)),"_")
		If QTemp(0) = title Then
			If left(QTemp(1),5) = "title" Then subjectTemp(i-1,1) = right(QTemp(1),len(QTemp(1))-6)
			If right(QTemp(1),6) = "single" Then subjectTemp(i-1,0) = "single"
			If right(QTemp(1),8) = "multiple" Then subjectTemp(i-1,0) = "multiple"
			If right(QTemp(1),8) = "question" Then subjectTemp(i-1,0) = "question"
			If left(QTemp(1),6) = "answer" Then
				subjectTemp(i-1,jj) = right(QTemp(1),len(QTemp(1))-8)
				jj = jj + 1
			End If
		End If
	Next
Next	

'ͨ�������ѭ�����Խ��б���ת������ΪRequest������������ת��
For i = 0 to QNum-1
	subject(i,1) = request.Form("Q" & (i+1) & "_title")
	subject(i,0) = request.Form("Q" & (i+1) & "_type")
	For j = 2 to 10
		If subjectTemp(i,j-1) <> "" Then		
			subject(i,j) = request.Form("Q" & (i+1) & "_answer" & (j-1))
		End If
	Next
Next

'���¿�ʼ�������ݿ����
On Error Resume Next

'������A�ӵ�ǰ�����ʾ��������ı������洢���������Ա�ύ�ĵ����ʾ�Ĵ���Ϣ
tableAttribute = "(id autoincrement primary key, Stu_ID varchar(20)"
For i = 1 to QNum
	tableAttribute = tableAttribute & " ,Q" & i & " varchar(200)"
Next
tableAttribute = tableAttribute & ")"
CreatTable = "create table A" & QN_ID & tableAttribute		'���ɴ������SQL���
conn.execute(CreatTable)	'������A�ӵ����ʾ��������ı������洢���������Ա�ύ�ĵ�����

'������Q�ӵ�ǰ�����ʾ��������ı������洢�����ʾ����Ϣ
tableAttribute = "(id autoincrement primary key, QN_ID varchar(20), QNum int, QType varchar(20),Option_Num int, QTitle varchar(100), Op1 varchar(200), Op2 varchar(200), Op3 varchar(200), Op4 varchar(200), Op5 varchar(200), Op6 varchar(200), Op7 varchar(200), Op8 varchar(200), Op9 varchar(200), Question varchar(200))"
CreatTable = "create table Q" & QN_ID & tableAttribute		'���ɴ������SQL���
conn.execute(CreatTable)	'������Q�ӵ����ʾ��������ı�



If Err.Number = 0 Then
	'����QN_Title���ñ�洢�����ʾ�ı�ţ���Ŀ��������������
	Set rsQuestion = Server.CreateObject("ADODB.RecordSet")
	sqlStrQN_Title = "select * from QN_Title"
	rsQuestion.open sqlStrQN_Title,conn,1,3	'�򿪵����ʾ�����ύ�����ʾ�ı�ţ����ƣ�����
	rsQuestion.addnew
	rsQuestion("QNum") = QNum
	rsQuestion("Title") = QTitle
	rsQuestion("QN_ID") = QN_ID
	rsQuestion("QVisible") = "1"
	rsQuestion.update		'���µ����ʾ��
	rsQuestion.close
	Set rsQuestion = nothing
	
	'�򿪵�ǰ�½��ĸ��ʾ�ı�
	Set rs = Server.CreateObject("ADODB.RecordSet")
	sqlStr = "select * from Q" & QN_ID
	rs.open sqlStr,conn,1,3

	For p = 0 to QNum-1
		Select Case subject(p,0)
			Case "single"
				rs.addnew
				ttt = 0
				rs("QN_ID") = QN_ID
				rs("QNum") = p + 1	'���
				rs("QType") = "single"
				rs("QTitle") = subject(p,1)
				For q1 = 2 to 10
					If subject(p,q1) <> "" Then
						rs("Op" & (q1-1)) = subject(p,q1)
						ttt = ttt + 1
					End If
				Next
				rs("Option_Num") = ttt	'�����ѡ����
				rs.update
			Case "multiple"
				rs.addnew
				ttt = 0
				rs("QN_ID") = QN_ID
				rs("QNum") = p + 1	'���
				rs("QType") = "multiple"
				rs("QTitle") = subject(p,1)
				For q2 = 2 to 10
					If subject(p,q2) <> "" Then
						rs("Op" & (q2-1)) = subject(p,q2)
						ttt = ttt + 1
					End If
				Next
				rs("Option_Num") = ttt
				rs.update
			Case "question"
				rs.addnew
				rs("QN_ID") = QN_ID
				rs("QNum") = p + 1	'���
				rs("QType") = "question"
				rs("QTitle") = "������"
				rs("Question") = subject(p,1)
				rs.update
			Case Else
				response.Write("Error")
		End Select	
	Next
	response.Write("<a href='index.asp' target='_self'>��ӳɹ����������</a>")
Else
	response.Write("�����������ϵ�����д��Ա��")
End If
End If
rs.close
Set rs = nothing
conn.close
Set conn = nothing 
%>