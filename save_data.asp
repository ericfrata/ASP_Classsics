<%
	set conexao = server.CreateObject ("ADODB.Connection")
		conexao.Open "database","user","password" 'user and password from ODBC 

	'Get parameters 
	parameter_01 = request("parameter_01")
    parameter_02 = request("parameter_02")

	' create sql query to insert data, in this case it's using a procedure
	
    set objCommandoSalvar = Server.CreateObject("ADODB.command")

    objCommandoSalvar.ActiveConnection = conexao
    objCommandoSalvar.CommandText = "SPR_PROCEDURE"
    objCommandoSalvar.CommandType = 4 

    objCommandoSalvar.Parameters.Append objCommandoSalvar.CreateParameter ("@parameter_name_from_proc_01", 3, 1, 0, parameter_01)
    objCommandoSalvar.Parameters.Append objCommandoSalvar.CreateParameter ("@parameter_name_from_proc_02", 3, 1, 0, parameter_02)
    
	'execute command in database
    objCommandoSalvar.Execute
%>
	
	
	