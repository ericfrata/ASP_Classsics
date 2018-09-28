<%

set conexao = server.CreateObject ("ADODB.Connection")
	conexao.Open "database","user","password" 'user and password from ODBC 

    sqlReport = "SPR_Procedure_to_Export"
    fileName="Exemple_export"
	

set rsReport = conexao.Execute(sqlReport)	

	
		Response.ContentType = "application/vdn.ms-excel"
		
		Response.AddHeader "Content-Disposition","attachment; filename=" & fileName & ".xls"
		Response.AddHeader "Accept-Ranges", "bytes"
		
		'HTML header to fix any problems with characters
		
		response.write "<html><head><meta http-equiv='Content-Language' content='pt-BR' />"
		response.write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head><body>"
	
		'create table HTML to use in the excel
		
		Response.Write "<table border='1'>"
		Response.Write "<tr align='center'>"
		
		'First line is the name of fields 
		For each item in rsReport.Fields
			response.write "<td bgcolor='silver'><b>" & item.Name & "</b></td>"
		Next
		
		'while to get all the data from the procedure
		While Not rsReport.EOF
			response.write "<tr>"
		
			For each item in rsReport.Fields
				response.write "<td>" & rsReport(item.Name) & "</td>"
			Next
			
			response.write "</tr>"
			rsReport.MoveNext
		Wend
		
		response.write "</table></body></html>"
	

%>