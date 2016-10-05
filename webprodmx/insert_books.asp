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
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_conn_webprodmx_STRING
  MM_editTable = "dbo.tbl_Books"
  MM_editRedirectUrl = "categories3.asp"
  MM_fieldsStr  = "BookTitle|value|BookAuthorFirstName|value|BookAuthorLastName|value|BookPrice|value|BookISBN|value|BookPageCount|value|BookImage|value|CategoryID|value"
  MM_columnsStr = "BookTitle|',none,''|BookAuthorFirstName|',none,''|BookAuthorLastName|',none,''|BookPrice|none,none,NULL|BookISBN|',none,''|BookPageCount|none,none,NULL|BookImage|',none,''|CategoryID|none,none,NULL"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim rsCategories
Dim rsCategories_numRows

Set rsCategories = Server.CreateObject("ADODB.Recordset")
rsCategories.ActiveConnection = MM_conn_webprodmx_STRING
rsCategories.Source = "SELECT CategoryID, Category  FROM dbo.tbl_Categories  ORDER BY Category"
rsCategories.CursorType = 0
rsCategories.CursorLocation = 2
rsCategories.LockType = 1
rsCategories.Open()

rsCategories_numRows = 0
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Add book</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
Add Book
<form method="post" action="<%=MM_editAction%>" name="form1">
  <table>
    <tr valign="baseline">
      <td nowrap align="right">Book Title:</td>
      <td>
        <input type="text" name="BookTitle" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Author First Name:</td>
      <td>
        <input type="text" name="BookAuthorFirstName" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Author Last Name:</td>
      <td>
        <input type="text" name="BookAuthorLastName" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Price:</td>
      <td>
        <input type="text" name="BookPrice" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">ISBN:</td>
      <td>
        <input type="text" name="BookISBN" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Page Count:</td>
      <td>
        <input type="text" name="BookPageCount" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Image Name:</td>
      <td>
        <input type="text" name="BookImage" value="" size="32">
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Category:</td>
      <td>
        <select name="CategoryID">
          <%
While (NOT rsCategories.EOF)
%>
          <option value="<%=(rsCategories.Fields.Item("CategoryID").Value)%>"><%=(rsCategories.Fields.Item("Category").Value)%></option>
          <%
  rsCategories.MoveNext()
Wend
If (rsCategories.CursorType > 0) Then
  rsCategories.MoveFirst
Else
  rsCategories.Requery
End If
%>
        </select>
      </td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">&nbsp;</td>
      <td>
        <input type="submit" value="Insert record">
      </td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1">
</form>
<p>&nbsp;</p>
</body>
</html>
<%
rsCategories.Close()
Set rsCategories = Nothing
%>
