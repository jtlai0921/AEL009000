<%@ Language=VBScript %>
<% Option Explicit %>
<%
    Dim Name, Sex, School, Interest, Thought
    Name=Request("UserName")
    Select Case Request("UserSex")
        Case "Male":
            Sex="�k"
        Case "Female":
            Sex="�k"
    End Select
    School=Request("UserSchool")
    Interest=Request("UserInterest")
    Thought=Replace (Request("UserThought"), vbCrLf, "<BR>")
%> 
<HTML>
    <HEAD>
        <TITLE>�ӤH�p�ɮ׽T�{����</TITLE>
    </HEAD>
    <BODY BGCOLOR="#D1EFFC">
        <CENTER><IMG SRC="profile2.jpg"></CENTER>
        <P><FONT COLOR="Blue"><B><I><%=Name%></B></I></FONT>�A�z�n�I���±z��g�ӤH��ƪ�A�z��J����Ʀp�U�G</P>
        <P>�ʧO�G<FONT COLOR="Blue"><%=Sex%></FONT></P>
        <P>�̰��Ǿ��G<FONT COLOR="Blue"><%=School%></FONT></P>
        <P>�z�P���쪺���ʡG<FONT COLOR="Blue"><%=Interest%></FONT></P>
        <P>�z�﫢�魷���ݪk�G<FONT COLOR="Blue"><BLOCKQUOTE><%=Thought%></BLOCKQUOTE></FONT></P>
    </BODY> 
</HTML>
	     