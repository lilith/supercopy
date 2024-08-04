//Reads a log file, and creates and colored, highlighted html file from it.
//Returns HTML representing a summary and link to the full version
function FormatOutput(filename,newfilename){
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	//var sh = new ActiveXObject("WScript.Shell");
	//var pf = sh.ExpandEnvironmentStrings("%PROGRAMFILES%");
	if (fso.FileExists(filename)){
	    ts = fso.OpenTextFile(filename, 1);
	    var linen = 0;
	    var insertstring = new Array();
		while (!ts.AtEndOfStream){
			insertstring[linen]=ts.ReadLine();
			linen++;
		}
		ts.Close();
		var styles = 
			"p.title1{margin-top: 0px; margin-bottom: 0px; FONT: caption; COLOR: silver; }" + 
			
			"p.title2{margin-top: 0px; margin-bottom: 0px;FONT-SIZE: large; COLOR: black; " + 
				"FONT-FAMILY: 'Times New Roman', Serif;BACKGROUND-COLOR: lightskyblue}" + 
				
			"p.command{margin-top: 20px; margin-bottom: 10px; font-family: Courier New; background-color:Transparent;" + 
				"font-size:large; font-weight:bold; color: Red; border-style: solid; border-width: thin; border-color: Black;}"+ 
						
			"p.result{margin-top: 6px; margin-bottom: 6px; font-family: Arial; background-color:Yellow;"+ 
				"font-size:large; font-weight:bold; color:Blue;}"+ 
				
			"p.error{margin-top: 25px; margin-bottom: 25px; font-family: Courier New; background-color: Red;"+ 
				"font-size:x-large; color: Black;}" + 
				
			"p.output{    margin-top: 0px; margin-bottom: 0px; font-family: Arial; background-color: Transparent;"+ 
				" font-size: small; color:Black;}";	
			
			//<link rel=\"stylesheet\" type=\"text/css\"href=\"Test.css\" />
		var string1 = "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\"><html><head><title>";
		var string2 = "</title><style>" + styles + "</style></head><body>" + 
						"<P class=\"title1\">XCOPY User Interface - Dynamically generated batch file" + 
						"</P><P class=\"title2\">XCOPY Results:</P>";

		var string3 = "<P class=\"output\">";
		var string4 = "</P></body></html>";
				
		ts = fso.OpenTextFile(newfilename, 2,true);
		ts.WriteLine(string1 + "XCopy User Interface" + string2);
		
		var commands = 0;
		var results = 0;
		var errors = 0;
		var cmdlist = "";
		var resultlist = "";
		var errorlist = "";
		var summary = ""
		
		
		var lineindex;
		for (lineindex = 0; lineindex < insertstring.length;lineindex++)
		{
			if ((insertstring[lineindex].substring(0,5)=="XCOPY") || 
				(insertstring[lineindex].substring(0,5)=="RMDIR") ||
				(insertstring[lineindex].substring(0,5)=="MKDIR") ||
				(insertstring[lineindex].substring(0,5)=="PAUSE") ||
				(insertstring[lineindex].substring(0,5)=="These") ||
				(insertstring[lineindex].substring(0,3)=="DEL")){
				
				ts.WriteLine("<P class=\"command\">"  + insertstring[lineindex] + "</P>");
				cmdlist += "<P class=\"command\">"  + insertstring[lineindex] + "</P>";
				summary += "<P class=\"command\">"  + insertstring[lineindex] + "</P>";
				commands++;
			}else if ((insertstring[lineindex].substr(insertstring[lineindex].length - 6,6)=="copied") ||
			(insertstring[lineindex].substr(insertstring[lineindex].length - 6,6)=="ile(s)")){
			
				ts.WriteLine("<P class=\"result\">"	 + insertstring[lineindex] + "</P>");
				resultlist +="<P class=\"result\">"	 + insertstring[lineindex] + "</P>"
				summary +="<P class=\"result\">"	 + insertstring[lineindex] + "</P>"
				results++;
			}else if ((insertstring[lineindex].substr(insertstring[lineindex].length - 7,7)=="enied. ") ||
					 (insertstring[lineindex].substr(insertstring[lineindex].length - 7,6)=="failed") || 
					 (insertstring[lineindex].substr(0,12)=="Insufficient") ||
					 (insertstring[lineindex].substr(0,3)=="The") ||
					 (insertstring[lineindex].substr(0,7)=="Invalid") ||
					 (insertstring[lineindex].substr(0,7)=="Sharing") ||
					 (insertstring[lineindex].substr(0,6)=="Cannot") ||
					 (insertstring[lineindex].substr(0,5)=="Can't") ||
					 (insertstring[lineindex].substr(0,6)=="Failed") ||
					 (insertstring[lineindex].substr(0,8)=="File not") ||
			         	 (insertstring[lineindex].substr(0,6)=="Access") ||
					 (
					(insertstring[lineindex].length == 3 ||
					insertstring[lineindex].length == 4) &&
					insertstring[lineindex].substr(0,1)=="[")){

			        if (insertstring[lineindex].length == 3 || insertstring[lineindex].length == 4){
					if (insertstring[lineindex].substr(1,1) != "0"){
						var Explanation = "Error Returned: " + insertstring[lineindex];
						var errornumber = insertstring[lineindex].substr(1,1);
						if (errornumber == "1") Explanation = "Error 1: No files were found to copy";
						if (errornumber == "2") Explanation = "Error 2: The user cancelled the copy with CTRL-C!";
						if (errornumber == "3") Explanation = "ErrorLevel 3: Unknown Error occured in XCOPY!";
						if (errornumber == "4") Explanation = "Error 4: Invalid drive name, syntax, or there is not enough memory or disk space! (initialization error) This is sometimes due to filesystem limitations. If you are using two drives, please make sure they both are using the same filesystem. FAT32 does not support files over 4 GB.";
						if (errornumber == "5") Explanation = "Error 5: Disk write error";
						ts.WriteLine("<P class=\"error\">" + Explanation + "</P>");
						errorlist += "<P class=\"error\">" + Explanation + "</P>";
						summary += "<P class=\"error\">" + Explanation + "</P>";
						errors++;
					}

				}else{ 
					ts.WriteLine("<P class=\"error\">" + insertstring[lineindex] + "</P>");
					errorlist += "<P class=\"error\">" + insertstring[lineindex] + "</P>";
					summary += "<P class=\"error\">" + insertstring[lineindex] + "</P>";
					errors++;
				}
			}else{
				ts.WriteLine("<P class=\"output\">" + insertstring[lineindex] + "</P>");
			}
			
		}
		
		if (errors > 0) summary += "<P>Note: Any errors displayed apply to the preceding operation. Errors are displayed below result.</P>";
		
		ts.WriteLine(string4);
		ts.Close();
		
		var url = newfilename;
		for (i=0;i<url.length;i++){
			if (url.substr(i,1)=="\\"){
				url = url.substr(0,i) + "/" + url.substr(i + 1,url.length - i -1);
			}
		}
		
		
		var CompleteButton = "<INPUT id=\"button1\" name=\"button2\" " +
		"type=\"button\" value=\"View Complete Listing\" language=\"javascript\" " +
		"onclick=\"window.open('" 
		+ url + "'); \" >";
		
		CompleteButton="<a href=\"" + url + "\">View complete result listing</a>";
		
		
		return ("<P class=\"infoline\"> "+ commands +
		 " command(s), " + 
		results + " result(s), " + errors + " errors." + "&nbsp;&nbsp;&nbsp;" + CompleteButton  +  "</p>" + summary);

	}
}
