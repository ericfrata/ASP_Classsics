
<html>
<head>
    <title>:: Eric Frata ::</title>
    <meta http-equiv="Content-Language" content="pt-BR" />
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    

    <script type="text/javascript" src="js/jquery-2.1.4.min.js"></script>
    <script type="text/javascript" src="js/bootstrap.min.js"></script>
    <link href='CSS/bootstrap.min.css' rel='stylesheet' type='text/css'>
    <link href='css/bgc_Style.css' rel='stylesheet' type='text/css'>
	<link rel="stylesheet" href="font-awesome/css/font-awesome.min.css">
    <!--[if lt IE 9]>
        <script src="../../js/html5shiv.min.js" type="text/javascript"></script>
        <script src="../../js/respond.min.js" type="text/javascript"></script>
    <![endif]-->

<%

set conexao = server.CreateObject ("ADODB.Connection")
	'odbc conection 
	conexao.Open "database","login","password"
	
	script = request("Script")
	
	
	if script = "" then
		'test script
		script = "select * from msdb.sys.data_spaces"
	end if
	
	set rsExecutar = server.CreateObject("ADODB.Recordset")
	set rsExecutar = conexao.Execute(script)
	
%>
	

</head>
<body>
<script type="text/javascript">

var ie = 0;
try { ie = navigator.userAgent.match( /(MSIE |Trident.*rv[ :])([0-9]+)/ )[ 2 ]; }
catch(e){}
	
if ((parseInt(ie) < 9) && ie != 0 ){
	alert("Para esta página funcionar corretamente favor desabilitar o modo de compatibilidade do navegador. \nCaso sua versão do Internet Explorer for anterior a versão 9, favor atualizar o mesmo.");
};

function Enviar() {
			document.form1.action = "acesso_banco.asp";
            document.form1.submit();
        }

</script>


    <div class="container">
	
        <p/>
  
        <div class="hero-unit">
            <h1>Acesso database</h1>
			
			
        </div>
		</p>
		<div class="borderDiv">
			<label for="Script">Script:</label>
			<form method="POST" id="form1" name="form1" >
			<textarea class="form-control" rows="5" id="Script" name="Script"></textarea>
			</p>
			<p align='center'>
			<a href="javascript:Enviar()" class="btn btn-primary btn-large">Executar</a>
			</p>
			
		</div>
		<p>
		<div class="borderDiv">
		
		
			 Executado o script:
        <br>
        " <%=script%> "<p>
		
        <p>
		</div>
		<p>
		<div>
		Retorno :<p><p><p>
            <%if not rsExecutar.EOF then%>
      
            <table class="table" border="1px" cellpadding="5px" cellspacing="0" ID="tabela">
                <%
					For each item in rsExecutar.Fields
					 response.write "<th>" & item.Name & "</th>"
					Next
                %>

				<%
					While Not rsExecutar.EOF
				%>
				<tr>
					 <%For each item in rsExecutar.Fields%>
							<td><%=rsExecutar(item.Name)%></td>
						<%Next%>
				</tr>
				<%
					rsExecutar.MoveNext
					Wend
				%>

                <%end if%>
            </table>
		
		
		</div>
		
		
		</p>
    </div>
</body>
</html>
<%
conexao.Close
%>