# ClassicASP.NUglify
This is a Component Object Model (COM) Dynamic-link library (DLL) coded in C# that can be set in Classic ASP using VBscripts "CreateObject" method and allows you to minify and compress JS and CSS code in Classic ASP.

Uses NUglify (a fork of the Microsoft Ajax Minifier)
https://github.com/trullock/NUglify

NUglify is a .NET library that provides minification for JavaScript and CSS. It supports modern JavaScript standards (ES2020 and ES2021) and offers features like removing comments, collapsing whitespaces, and compressing inline styles and scripts. 

This can be very useful for minifying, compressing and generally tidying up large JavaScript or CSS code generated programmatically in Classic ASP.

## INSTALLATION:

Make sure you have the lastest .NET Framework installed (tested on .NET Framework 4.7.2)
	
Open IIS, go to the applicaton pools and open the pool being used by your 
Classic ASP app. Check the .NET CRL version
E.g: v4.0.30319
	
Navigate to the CRL folder
E.g: C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
Copy over: ClassicASP.NUglify.dll and NUglify.dll
	
Run CMD as administrator

Change the directory to your CRL folder
E.g: cd C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
Run the following command: RegAsm ClassicASP.NUglify.dll /tlb /codebase

Restart IIS

## Usage

### Minify and Compress JS/CSS:


	Dim NUglify : Set NUglify = Server.CreateObject("ClassicASP.NUglify")

		NUglify.JS("JS Code")
		
		NUglify.CSS("CSS Code")

	Set NUglify = Nothing

## test/NUglify.asp results:

### NUglify, bootstrap.js Test:

Time to NUglify: 0.0156s

Original Length: 145,401

NUglify Length: 61,726

Compression Ratio: 57.55%

-----

### NUglify, bootstrap.css Test:

Time to NUglify: 0.0469s

Original Length: 281,046

NUglify Length: 232,145

Compression Ratio: 17.4%
