<%@ Language=VBScript %>
<% Option Explicit %>
<%
    Dim Name, Sex, School, Interest, Thought
    Name=Request("UserName")
    Select Case Request("UserSex")
        Case "Male":
            Sex="男"
        Case "Female":
            Sex="女"
    End Select
    School=Request("UserSchool")
    Interest=Request("UserInterest")
    Thought=Replace (Request("UserThought"), vbCrLf, "<BR>")
%> 
<HTML>
    <HEAD>
        <TITLE>個人小檔案確認網頁</TITLE>
    </HEAD>
    <BODY BGCOLOR="#D1EFFC">
        <CENTER><IMG SRC="profile2.jpg"></CENTER>
        <P><FONT COLOR="Blue"><B><I><%=Name%></B></I></FONT>，您好！謝謝您填寫個人資料表，您輸入的資料如下：</P>
        <P>性別：<FONT COLOR="Blue"><%=Sex%></FONT></P>
        <P>最高學歷：<FONT COLOR="Blue"><%=School%></FONT></P>
        <P>您感興趣的活動：<FONT COLOR="Blue"><%=Interest%></FONT></P>
        <P>您對哈日風的看法：<FONT COLOR="Blue"><BLOCKQUOTE><%=Thought%></BLOCKQUOTE></FONT></P>
    </BODY> 
</HTML>
	     