<%@ Language=VBScript %>
<% Option Explicit %>
<%
    Dim Name, Mail, Message

    Name=Request("UserName")
    Mail=Request("UserMail")
    Message=Replace (Request("UserMessage"), vbCrLf, "<BR>")
%>
<HTML>
    <HEAD>
        <TITLE>��ܯd��</TITLE>
    </HEAD>
    <BODY>
        <TABLE WIDTH="100%">
            <TR><TD>�m�W&nbsp;&nbsp;�G</TD></TR>
            <TR><TD BGCOLOR="#FBE0FB"><%=Name%></TD></TR>  
            <TR><TD>E-Mail�G</TD></TR>
            <TR><TD BGCOLOR="#CBFCDA"><%=Mail%></TD></TR>
            <TR><TD>�d��&nbsp;&nbsp;�G</TD></TR>
            <TR><TD BGCOLOR="#FFFFC8"><%=Message%></TD></TR>
        </TABLE>
    </BODY> 
</HTML>
	     