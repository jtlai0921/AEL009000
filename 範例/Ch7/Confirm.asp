<%@ Language=VBScript %>
<% Option Explicit %>
<%
    Dim Name, Mail, Age, Phone, Trouble, Number
    Name=Request("UserName")
    Mail=Request("UserMail")
    Select Case Request("UserAge")
        Case "Age1":
            Age="����20��"
        Case "Age2":
            Age="20~29"
        Case "Age3":
            Age="30~39"
        Case "Age4":
            Age="40~49"
        Case "Age5":
            Age="50���H�W"
    End Select
    Phone=Request("UserPhone")
    Trouble=Request("UserTrouble")
    Number=Request("UserNumber")
%> 
<HTML>
    <HEAD>
        <TITLE>�j���j�ϥηN���լd��T�{����</TITLE>
    </HEAD>
    <BODY BACKGROUND="free0.gif">
        <P><IMG SRC="free1.jpg"></P>
        <P><I><%=Name%></I>�A�z�n�I���±z��g�N���լd��A�z��J����Ʀp�U�G</P>
        �q�l�l��a�}�G<%=Mail%><BR>
        �~�֡G<%=Age%><BR>
        ���g�ϥιL������t�P�G<%=Phone%><BR>
        �ϥΤ���ɳ̱`�I�쪺���D�G<%=Trouble%><BR>
        �ϥέ��a�~�̪������G<%=Number%>
    </BODY> 
</HTML>
	     