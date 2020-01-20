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
        <TITLE>顯示留言</TITLE>
    </HEAD>
    <BODY>
        <TABLE WIDTH="100%">
            <TR><TD>姓名&nbsp;&nbsp;：</TD></TR>
            <TR><TD BGCOLOR="#FBE0FB"><%=Name%></TD></TR>  
            <TR><TD>E-Mail：</TD></TR>
            <TR><TD BGCOLOR="#CBFCDA"><%=Mail%></TD></TR>
            <TR><TD>留言&nbsp;&nbsp;：</TD></TR>
            <TR><TD BGCOLOR="#FFFFC8"><%=Message%></TD></TR>
        </TABLE>
    </BODY> 
</HTML>
	     