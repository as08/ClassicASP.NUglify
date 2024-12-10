<%
	
	' INSTALLATION:
	'************************************************************************
	
	' Uses NUglify (a fork of the Microsoft Ajax Minifier)
	' https://github.com/trullock/NUglify
	
	' Make sure you have the lastest .NET Framework installed (tested on .NET Framework 4.7.2)
	
	' Open IIS, go to the applicaton pools and open the application pool being used by your 
	' website. Check the .NET CRL version
	' E.g: v4.0.30319
	
	' Navigate to the CRL folder
	' E.g: C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
	' Copy over: 
	' ClassicASP.NUglify.dll
	' NUglify.dll
	
	' Run CMD as administrator
	
	' Change the directory to your CRL folder
	' E.g: cd C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
	' Run the following command: RegAsm ClassicASP.NUglify.dll /tlb /codebase
	
	
	' NUglify IN VBSCRIPT:
	'************************************************************************
	
	Function ReadFile(filePath)
	
		Dim fso, file, fileContent
	
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
	
		If fso.FileExists(filePath) Then
			' Open the file and read its content
			Set file = fso.OpenTextFile(filePath, 1) ' 1 = ForReading
			fileContent = file.ReadAll
			file.Close
			Set file = Nothing
		Else
			' If the file doesn't exist, return an error message
			fileContent = "File not found: " & filePath
		End If
	
		Set fso = Nothing
	
		ReadFile = fileContent
		
	End Function
	
	Dim NUglify, NUglified, start, bootstrap_js, bootstrap_css, origLen
	
	Set NUglify = Server.CreateObject("ClassicASP.NUglify")
		
		' JS test
		
		bootstrap_js = ReadFile(Server.MapPath("bootstrap.js"))
		origLen = Len(bootstrap_js)
		
		start = Timer()
		
		NUglified = NUglify.JS(bootstrap_js)
		
		Response.Write "<h1>NUglify, bootstrap.js Test:</h1><textarea rows=""30"" cols=""100"">" & NUglified & "</textarea>"
		
		Response.Write "<p><b>Time to NUglify:</b> " & formatNumber(Timer()-start,4) & "s</p>"
		Response.Write "<p><b>Original Length:</b> " & formatNumber(origLen,0) & "</p>"
		Response.Write "<p><b>NUglify Length:</b> " & formatNumber(Len(NUglified),0) & "</p>"
		Response.Write "<p><b>Compression Ratio:</b> " & Round((1-(Len(NUglified)/origLen))*100,2) & "%</p><hr>"
		
		' CSS test
		
		bootstrap_css = ReadFile(Server.MapPath("bootstrap.css"))
		origLen = Len(bootstrap_css)
		
		start = Timer()
		
		NUglified = NUglify.CSS(bootstrap_css)
		
		Response.Write "<h1>NUglify, bootstrap.css Test:</h1><textarea rows=""30"" cols=""100"">" & NUglified & "</textarea>"
		
		Response.Write "<p><b>Time to NUglify:</b> " & formatNumber(Timer()-start,4) & "s</p>"
		Response.Write "<p><b>Original Length:</b> " & formatNumber(origLen,0) & "</p>"
		Response.Write "<p><b>NUglify Length:</b> " & formatNumber(Len(NUglified),0) & "</p>"
		Response.Write "<p><b>Compression Ratio:</b> " & Round((1-(Len(NUglified)/origLen))*100,2) & "%</p>"
	
	Set NUglify = Nothing
		
%>