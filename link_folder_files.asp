

<html>
<head>
</head>
<body> 
<%
	Dim Local, Folder, File, ObjFS, objRootFolder

	Local = "C:\Inetpub\wwwroot\Site\Folder" 'Exemple folder

	'Create the FileSystemObject
	Set ObjFS = Server.CreateObject("Scripting.FileSystemObject")
	Set objFolder = ObjFS.GetFolder(Local)
%>
<!-- Create the table with column description -->
<table border='1' cellpadding=2 cellspacing=0 width='100%' style='font-family: Tahoma, Arial; font-size: 11px;'>
<tr>
<td><b>Nome</b></td>
<td><b>Data da última modificação</b></td>
</tr>

<%
'Loop to show the founded files 
For Each File in objFolder.files
	numCalc= Len(File.Name)
	Response.Write " <tr>"
	Response.Write " <td><a href='arquivos/" & File.Name & "' target='_blank'>" & Replace(left(File.Name,numCalc-4),"_"," ") & "</a></td>"
	Response.Write " <td>" & File.DateLastModified & "</td>"
	Response.Write " </tr>"
Next
%>

</table>
</body>
</html>