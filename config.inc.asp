<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit '//强制声明全局变量
Response.CodePage=65001
Response.Charset="UTF-8"
Server.ScriptTimeOut = 90
Response.Buffer = True
Response.CacheControl="no-cache"
%>
<!--#include file="include/config.asp"-->
<!--#include file="include/db.class.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="include/des.asp"-->