<%@language=VBScript%>
<%  Option Explicit
    Response.Expires=0
%>

<!-- #include file = "include/adovbs.asp" -->
<!-- #include file = "include/startup.asp" -->
<!-- #include file = "include/SrvrFunctions.asp" -->
<!-- #include file = "include/EventFunctions1.asp" -->

<!-- #include file = "RenderFunctions.asp" -->
<!-- #include file = "ClassActivityForm.asp" -->


<%
dim rst, rst2, sSelected
set rst = CreateObject("ADODB.RecordSet")
set rst2 = CreateObject("ADODB.RecordSet")

rst.CursorLocation = adUseClient
rst2.CursorLocation = adUseClient

RecoverSQLConnection()

RecoverSession(true) 



dim actForm
set actForm = loadActivityForm(40)

%>

actForm.idActivity = [<%=actForm.idActivity %>]<br />
actForm.idForm = [<%=actForm.idForm %>]<br />
actForm.numResponses = [<%=actForm.numResponses %>]<br />


<%
dim r
for each r in actForm.responses %>
    idQ = [<%=r.idQuestion %>] 
    idR = [<%=r.idResponse %>] <br />
<%next%>