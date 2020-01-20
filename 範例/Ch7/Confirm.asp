<%@ Language=VBScript %>
<% Option Explicit %>
<%
    Dim Name, Mail, Age, Phone, Trouble, Number
    Name=Request("UserName")
    Mail=Request("UserMail")
    Select Case Request("UserAge")
        Case "Age1":
            Age="未滿20歲"
        Case "Age2":
            Age="20~29"
        Case "Age3":
            Age="30~39"
        Case "Age4":
            Age="40~49"
        Case "Age5":
            Age="50歲以上"
    End Select
    Phone=Request("UserPhone")
    Trouble=Request("UserTrouble")
    Number=Request("UserNumber")
%> 
<HTML>
    <HEAD>
        <TITLE>大哥大使用意見調查表確認網頁</TITLE>
    </HEAD>
    <BODY BACKGROUND="free0.gif">
        <P><IMG SRC="free1.jpg"></P>
        <P><I><%=Name%></I>，您好！謝謝您填寫意見調查表，您輸入的資料如下：</P>
        電子郵件地址：<%=Mail%><BR>
        年齡：<%=Age%><BR>
        曾經使用過的手機廠牌：<%=Phone%><BR>
        使用手機時最常碰到的問題：<%=Trouble%><BR>
        使用哪家業者的門號：<%=Number%>
    </BODY> 
</HTML>
	     