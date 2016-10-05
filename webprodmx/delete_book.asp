<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/conn_webprodmx.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="Administrator"
MM_authFailedURL="login.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (false Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%

if(Request.QueryString("BookID") <> "") then cmdDeleteBook__varBookID = Request.QueryString("BookID")

%>
<%

set cmdDeleteBook = Server.CreateObject("ADODB.Command")
cmdDeleteBook.ActiveConnection = MM_conn_webprodmx_STRING
cmdDeleteBook.CommandText = "DELETE FROM dbo.tbl_Books  WHERE BookID = " + Replace(cmdDeleteBook__varBookID, "'", "''") + ""
cmdDeleteBook.CommandType = 1
cmdDeleteBook.CommandTimeout = 0
cmdDeleteBook.Prepared = true
cmdDeleteBook.Execute()

%>
<%
Response.Redirect("categories3.asp")
%>