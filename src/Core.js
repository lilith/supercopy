//Constants
//Old temp dir:%SYSTEMROOT%\\TEMP
var LogFileValuepm = "%TEMP%\\%DATE:~-2,2%%DATE:~-10,2%%DATE:~-7,2%%Time:~0,2%%Time:~3,2%%TIME:~6,2%.supercopy.log";
var LogFileValueam = "%TEMP%\\%DATE:~-2,2%%DATE:~-10,2%%DATE:~-7,2%0%Time:~1,1%%Time:~3,2%%TIME:~6,2%.supercopy.log";
var LogFileVar = "%xcopylog%";
var LogFileName = "xcopylog";
var ExcludeFileValuepm = "%TEMP%\\%DATE:~-2,2%%DATE:~-10,2%%DATE:~-7,2%%Time:~0,2%%Time:~3,2%%TIME:~6,2%";
var ExcludeFileValueam = "%TEMP%\\%DATE:~-2,2%%DATE:~-10,2%%DATE:~-7,2%0%Time:~1,1%%Time:~3,2%%TIME:~6,2%";
var ampm = "%Time:~0,1%";
var ExcludeFileVar = "%ExcludedFiles%";
var ExcludeFileName = "ExcludedFiles";

//Parses a multi-line string and returns an array of the lines
function GetList(MultilineString){
    l("+GetList(\"" + MultilineString + "\") //Parses a multi-line string and returns an array of the lines");
	var gflfso
	gflfso = new ActiveXObject("Scripting.FileSystemObject");
	if (MultilineString == null) return null;
	//Source List
	var FileList;
	FileList = MultilineString.split("\n");
	//The list returned by the function. Only significant entries will be added to this one
	var FinishedList = new Array();
	//How far we are in the finished list array. This is diffferent from SourceIndex in that it is only incremented on a
	//significant (non-blank or whitespace) item
	var TargetIndex = 0;
	for ( SourceIndex = 0; SourceIndex < FileList.length; SourceIndex++ ){
	    //Trim off spaces and line returns
		FileList[SourceIndex]= Trim(FileList[SourceIndex]);
		if (FileList[SourceIndex].length > 0) {
			FinishedList[TargetIndex]=FileList[SourceIndex];
			TargetIndex++;
		}
	}
	l("-");
	return FinishedList;
}

//Trim off spaces and line returns from the beginning and end
function Trim(str){
	var newstr = str;
	var i;
	var offset = 0;
	for (i = 0; i < newstr.length;i++){
		if (newstr.charCodeAt(i + offset)==13){
			newstr = newstr.substring(1,newstr.length);
			offset--;
		}else if (newstr.charCodeAt(i + offset)==10){
			newstr = newstr.substring(1,newstr.length);
			offset--;
		}else if (newstr.charCodeAt(i + offset)==32){
			newstr = newstr.substring(1,newstr.length);
			offset--;
		}else{break;}
	}
	for (i = newstr.length - 1; i >= 0;i--){
		if (newstr.charCodeAt(i)==13){
			newstr = newstr.substring(0,newstr.length -1);
		}else if (newstr.charCodeAt(i)==10){
			newstr = newstr.substring(0,newstr.length -1);
		}else if (newstr.charCodeAt(i)==32){
			newstr = newstr.substring(0,newstr.length -1);
		}else{break;}
	}
	return newstr;
}


//Parses a multi-line file list and returns an array of files and folders.
//This is different from GetList in that it runs fso.GetAbsolutePathName on each array member
function GetFileList(ms){
    l("+GetFileList(\"" + ms + "\") //Parses a multi-line string and returns an array of the lines, after running each through fso.GetAbsolutePathName");
	var gflfso
	gflfso = new ActiveXObject("Scripting.FileSystemObject");
	if (ms == null) return null;
	var FileList;
	FileList = ms.split("\n");
	var FinishedList = new Array();
	var TargetIndex = 0;
	for ( SourceIndex = 0; SourceIndex < FileList.length; SourceIndex++ ){
		if (FileList[SourceIndex].charCodeAt(FileList[SourceIndex].length - 1) == 13){
			FileList[SourceIndex] = FileList[SourceIndex].substring(0,FileList[SourceIndex].length -1);
		}
		if (FileList[SourceIndex].length > 0) {
			FinishedList[TargetIndex]=gflfso.GetAbsolutePathName(FileList[SourceIndex]);
			TargetIndex++;
		}
	}
	l("-");
	return FinishedList;
}


//This function is in charge of creating the xcopy, mkdir, and rmdir commands for SuperCopy
//In both testing and real mode. 
//Sourcefiles should be an array of valid filenames to transfer
//exclusionfiles should be an array of full or partial, but absolute filenames to exclude. folders like \obj\ or c:\Bin
//Destdir should be a valid string path of a directry
//testonly and pipefile are used for projecting the results of an action.
//for testing, set testonly to true and pipefile to the file to read for the results.
//useexclusion turns on the /EXCUDE tag and uses the exclusionfiles argument.
//logoutput and pipefile can be used to create a log file.
//fordisplay, when true, formats output for command listing on-screen. It removes non-functional echo statements
//addpause Adds a pause command to the end of batch files
//GetCommands(string array,string array,string,bool,bool,string,bool,bool)
function GetCommands(sourcefiles,exclusionfiles,destdir,testonly,logoutput,pipefile,useexclusion,fordisplay,addpause){
    l("+GetCommands(sourcefiles,exclusionfiles,\"" + destdir + "\", testonly = " + testonly + ", logoutput = " + logoutput + ", pipefile = " + pipefile + 
    ", useexclusion = " + useexclusion + ", fordisplay = " + fordisplay + ", addpause = " + addpause + ")");
    
	//To store commands in during their compilation
	var commands = new Array();
	
	//To keep track of the current command index
	var cindex = 0;

	//Compute XCOPY arguments based upon user selections
	var arglist = GetArgs();
	//If testing, add the "try only /L" argument to the xcopy argument list
	if (testonly) 
		//Only add it if it isn't already added
		/*if (!chkTest.checked)*/
		 arglist += " /L";
		
	// FORCE + SILENT (No Questions) + Only affect archived files?AA
	//Compute DELETE arguments based upon 'limit to archive files' selection 
	// (rmdir still deletes all, clean will result in all archive files in directories.
	var delargs = " /F /Q";
	if (chkArchiveTag.checked)		{delargs += " /AA";}
	if (chkArchiveOnly.checked)		{delargs += " /AA";}	
	//Create file system object
	var sfso;
	sfso = new ActiveXObject("Scripting.FileSystemObject");
	
	//If this is not for display, we need to keep a copy of the original settings so that we can open it later.
	if (!fordisplay){
	    //Header information
    	
	    //This information will be used to read the configuration of this transfer 
	    //back into QXCopy during a file open
	    commands[cindex]="@REM Source Files";
	    cindex++;
	    for (var i = 0; i < sourcefiles.length; i++){
		    commands[cindex]="@REM " + sourcefiles[i];
		    cindex++;
	    }
	    commands[cindex]="@REM Exclusion List";
	    cindex++;
	    for (var i = 0; i < exclusionfiles.length; i++){
		    commands[cindex]="@REM " + exclusionfiles[i];
		    cindex++;
	    }
	    commands[cindex]="@REM Destination: " + destdir;
	    cindex++;
    	
	    var cfgstr = "";
	    StoreSettings();
	    var sitems = (new VBArray(settings.Items())).toArray();
	    for (var j=0;j < sitems.length;j++){
			    cfgstr += sitems[j] + " ";
	    }	
	    commands[cindex]="@REM Configuration: " + cfgstr;
	    cindex++;
    	
	}
	
	if (!fordisplay){
        //Formatting
	    commands[cindex]="@echo off";
	    cindex++;
	    commands[cindex]="COLOR F0";
	    cindex++;
	    commands[cindex]="echo %TIME% %DATE%";
	    cindex++;
	}

	//Create a Log File Name Var
	if (!testonly && logoutput){
	    commands[cindex]="REM Create a unique log filename that will sort correctly";cindex++;
		commands[cindex]="IF \"" + ampm + "\"==\" \" (";cindex++;
		commands[cindex]="   Set " + LogFileName + "=" + LogFileValueam;cindex++;
		commands[cindex]=") ELSE (";cindex++;
		commands[cindex]="   Set " + LogFileName + "=" + LogFileValuepm;cindex++;
		commands[cindex]=")";cindex++;
	}
	if (useexclusion){
	    commands[cindex]="REM Create a unique exclusion list filename";cindex++;
	    commands[cindex]="IF \"" + ampm + "\"==\" \" (";cindex++;
		commands[cindex]="   Set " + ExcludeFileName + "=" + ExcludeFileValueam;cindex++;
		commands[cindex]=") ELSE (";cindex++;
		commands[cindex]="   Set " + ExcludeFileName + "=" + ExcludeFileValuepm;cindex++;
		commands[cindex]=")";cindex++;
           
        if (!fordisplay){
		    //For debuggging
	        if (testonly || logoutput){
	            commands[cindex]="echo Excluded files and folders will not be included:" + " >> \"" +  pipefile + "\" 2>>&1";
		        cindex++;
		    }else{
			    commands[cindex]="echo Excluded files and folders will not be included:";
		        cindex++;
		    }
		}
		for (var i = 0; i < exclusionfiles.length; i++){
		    if (!fordisplay){
		        //For the benefit of debugging
		        if (testonly || logoutput){
		            //in both cases pipefile is what will be viewed
		            commands[cindex]="echo " + (i+1)  + ") " + exclusionfiles[i]  + " >> \"" +  pipefile + "\" 2>>&1";
		            cindex++;
		        }else{
		            //Here, echo it to screen
		            commands[cindex]="echo " + (i+1)  + ") " + exclusionfiles[i];
		            cindex++;
		        }
		    }
		    //real command
		    commands[cindex]="echo " + exclusionfiles[i]  + " >> \"" +  ExcludeFileVar + "\" 2>>&1";
		    cindex++;
	    }
	}

    var createdDir = false;
	//Check if we need to create the destination directory first.
	if (destdir.length == 0){
		if (!fordisplay){
			window.alert("Please specify a destination directory for this transfer");
			return null;
		}
	}else	if (sfso.FileExists(destdir)){
		window.alert("The destination directory you have specified is a file. Please specify a directory");
		return null;
	}else if (sfso.DriveExists(destdir)){
		destdir = sfso.GetDrive(destdir).Path;
		if (destdir.charAt(destdir.length - 1)==":") destdir +="\\";
		if (destdir.charAt(destdir.length - 1)!="\\") destdir +="\\";
	}else if (sfso.FolderExists(destdir)){
		destdir = sfso.GetFolder(destdir).Path;
		if (destdir.charAt(destdir.length - 1)!="\\") destdir +="\\";
		
	}else if (destdir.length > 0){
	
	    if (destdir.charAt(destdir.length - 1)!="\\") destdir +="\\";
	    
	    
		//window.alert(destdir); ???
		if (testonly || logoutput){
		    if (!fordisplay){
			    commands[cindex]="echo MKDIR  \"" + destdir + "\"" + " >> \"" + pipefile + "\" 2>>&1";
			    cindex++;
			}
			if (!testonly){
				commands[cindex]="MKDIR  \"" + destdir + "\"" + " >> \"" + pipefile + "\" 2>>&1";
				cindex++;
			}
		}else{
		    if (!fordisplay){
			    //In real mode, we both echo and execute
			    commands[cindex]="echo MKDIR \"" + destdir + "\"";
			    cindex++;
			}
			commands[cindex]="MKDIR \"" + destdir + "\"";
			cindex++;
		}
		
		
	}

	
	//Loop through each source file or directory
	for (i=0;i<sourcefiles.length;i++){
		var source; source = sourcefiles[i];
		var CreateDestSubfolder = false;
		var args;
		if (sfso.FolderExists(source) | sfso.DriveExists(source)){
			if (source.charAt(source.length - 1)!="\\") {
				if (source.length > 2) {
                    //Make sure it's not just a share
                    if (!sfso.DriveExists(source)){
				        CreateDestSubfolder = true;			
				    }	
				}
				source += "\\";
			}
			source += "*.*";
		}
		args = arglist;
		if (!sfso.FileExists(source)){
		    //Only use the Everything switch on Folders and wildcards, it is unpredictable on files
		    args += " /E"
		    if ((useexclusion) && (exclusionfiles.length > 0)){ 
		        //Add the Exclude switch as well, but only if the exclusion list isn't empty.
		        //Otherwise, XCOPY will throw an error.
		        args += " /EXCLUDE:" + ExcludeFileVar;
		    }
		}
		
		var usedel = false;
		
		//Use sourcefiles[i] here instead of source, so that we won't pick up the *.*
		var target; 
		if (destdir==""){
			target="";
		}else{
			if (CreateDestSubfolder){
				target = destdir + sfso.GetFolder(sourcefiles[i]).Name;
			}else if (source.indexOf("*")< 0 & source.indexOf("?") < 0){
				usedel=true;
				target = destdir + sfso.GetFileName(sourcefiles[i]);
			}else{
				usedel = true;
				target=destdir + source.substring(source.lastIndexOf("\\") + 1,source.length);
			}
		}
		
		//If we are doing a clean before an XCOPY, we must add del and rmdir arguments to affect target files
		if (optClean.checked){
			//For folders, use RMDIR, for files use DEL
			if (!usedel){
				//In testing mode we only echo the command
				//In testing mode, all output is piped to a file
				if (testonly || logoutput){
				    if (!fordisplay){
					    commands[cindex]="echo RMDIR /S /Q \"" + target + "\"" + " >> \"" + pipefile + "\" 2>>&1";
					    cindex++;
					}
					if (!testonly){
						commands[cindex]="RMDIR /S /Q \"" + target + "\"" + " >> \"" + pipefile + "\" 2>>&1";
						cindex++;
					}
				}else{
				    if (!fordisplay){
					    //In real mode, we both echo and execute
					    commands[cindex]="echo RMDIR /S /Q \"" + target + "\"";
					    cindex++;
					}
					commands[cindex]="RMDIR /S /Q \"" + target + "\"";
					cindex++;
				}
			}else{
				if (testonly || logoutput){
				    if (!fordisplay){
					    commands[cindex]="echo DEL \"" + target + "\"" + delargs + " >> \"" + pipefile+ "\" 2>>&1";
					    cindex++;
					}
					if (!testonly){
						commands[cindex]="DEL \"" + target + "\"" + delargs + " >> \"" + pipefile+ "\" 2>>&1";
						cindex++;
					}
				}else{
				    if (!fordisplay){
					    commands[cindex]="echo DEL \"" + target + "\"" + delargs;
					    cindex++;
					}
					commands[cindex]="DEL \"" + target + "\"" + delargs;
					cindex++;
				}
			}
		}
		//We can only use the prompt args in script mode, without logging
		var PromptArgs = "";
		if (chkPrompt.checked) PromptArgs = "/W";
		
		
		//CreateDestSubfolder occurs when we are copying a folder and its contents instead of just the contents
		if (CreateDestSubfolder){
			if (target.charAt(target.length - 1) != "\\") target += "\\";
			if (testonly || logoutput){
			        if (!fordisplay){
					    commands[cindex]="echo MKDIR \"" + target + "\"" + " >> \"" + pipefile + "\" 2>>&1";
					    cindex++;
					}
				if (!testonly){
					commands[cindex]="MKDIR \"" + target + "\"" + " >> \"" + pipefile + "\" 2>>&1";
					cindex++;	
				}
			}else{
				    //In real mode, we both echo and execute			
			    if (!fordisplay){
				    commands[cindex]="echo MKDIR \"" + target + "\"";
				    cindex++;
				}
				commands[cindex]="MKDIR \"" + target + "\"";
				cindex++;
			}
			if (testonly || logoutput){
			    if (!fordisplay){
				    commands[cindex]="echo XCOPY \"" + source + "\" \"" + target + "\"" + args + " >> \"" + pipefile+ "\" 2>>&1";
				    cindex++;
				}
				commands[cindex]="XCOPY \"" + source + "\" \"" + target + "\"" + args + " >> \"" + pipefile+ "\" 2>>&1";
				cindex++;
			}else{
				//Same as above, but not piping to a file
				if (!fordisplay){
				    commands[cindex]="echo XCOPY \"" + source + "\" \"" + target + "\"" + args+ " " + PromptArgs;
				    cindex++;
				}
				commands[cindex]="XCOPY \"" + source + "\" \"" + target + "\"" + args + " " + PromptArgs;
				cindex++;
			}
		}else{
			if (testonly || logoutput){
			    if (!fordisplay){
				    commands[cindex]="echo XCOPY \"" + source + "\" \"" + destdir + "\"" + args + " >> \"" + pipefile+ "\" 2>>&1";
				    cindex++;
				}
				commands[cindex]="XCOPY \"" + source + "\" \"" + destdir + "\"" + args + " >> \"" + pipefile+ "\" 2>>&1";
				cindex++;
			}else{
				//Same as above, but not piping to a file
				if (!fordisplay){
				    commands[cindex]="echo XCOPY \"" + source + "\" \"" + destdir + "\"" + args+ " " + PromptArgs;
				    cindex++;
				}
				commands[cindex]="XCOPY \"" + source + "\" \"" + destdir + "\"" + args+ " " + PromptArgs;
				cindex++;
			}
			//Now log the errorlevel
			if (testonly || logoutput){
				commands[cindex]="echo [%ERRORLEVEL%]" + " >> \"" + pipefile+ "\" 2>>&1";
				cindex++;
			}
		}
		
	}
	
	
	

	
	if (addpause){
		if (testonly || logoutput){
	        if (!fordisplay){
		        commands[cindex]="echo PAUSE" + " >> \"" + pipefile + "\" 2>>&1";
		        cindex++;
		    }
		    if (!testonly){
			    commands[cindex]="PAUSE";
			    cindex++;
		    }
	    }else{
	        if (!fordisplay){
		        //In real mode, we both echo and execute
		        commands[cindex]="echo PAUSE";
		        cindex++;
		    }
		    commands[cindex]="PAUSE";
		    cindex++;
	    }
	}
	l("-");
	return commands;
	
}

//Creates a summary of what will occur upon command execution
function GetSummary(){
	var GetSummaryVar = "\n - If you use wildcards, remember they apply to subfolder contents.";
	
	if (chkExclude.checked){
	    GetSummaryVar += "\n - Files and folders listing in the Exclude List";
	    GetSummaryVar += "\n - will not be copied. Patterns are also evaluated.";
	}
	
	if (chkLog.checked){
		GetSummaryVar += "\n - Output will be logged to a file. Execution will be ";
		GetSummaryVar += "\n - suspended indefinitely if input is required!";		
	}
	if (chkPrompt.checked){
		GetSummaryVar += "\n - Confirmation will be requested.";
	}
	
	if (chkArchiveTag.checked){
		GetSummaryVar += "\n - Only items tagged 'archive' will be copied.";
		GetSummaryVar += "\n - The 'archive' attribute will be copied to the destination files.";
		GetSummaryVar += "\n - The 'archive' attribute will be turned off on the source files.";
	}else if (chkArchiveOnly.checked){
		GetSummaryVar += "\n - Only items tagged 'archive' will be copied.";
		GetSummaryVar += "\n - The 'archive' attribute will be copied to the destination files.";
	}
	if (chkDecrypt.checked)
		GetSummaryVar += "\n - Items will be decrypted during the transfer.";
	if (chkHiddenSystem.checked)
		GetSummaryVar += "\n - Hidden files and System files will be copied.";
	if (chkResume.checked)
		GetSummaryVar += "\n - Copy will resume after a connection failure";
	if (chkMaintainReadOnly.checked){
		GetSummaryVar += "\n - The read-only attribute for files will be maintained.";
	}
	if (chkMaintainAudit.checked)
		GetSummaryVar += "\n - All file audit and security information will be maintained.";
	else if (chkMaintainSecurity.checked)
		GetSummaryVar += "\n - Security and file ownership information will be maintained.";

	//GetSummaryVar += "\n - Ignoring errors, piping not allowed.";
	
	if (optOverwrite.checked) 
		GetSummaryVar = "Source files will overwrite destination files." + GetSummaryVar;
	else if (optContinue.checked)
		GetSummaryVar = "Files are overwritten only if the source file is newer than the destination file. If the timestamp on both files is identical, nothing is done. Naturally, this type of copy always starts where it left off." + GetSummaryVar;
	else if( optContinue2.checked)
		GetSummaryVar = "Files are overwritten only if the source file is newer than the destination file. If the timestamp on both files is identical, nothing is done. Naturally, this type of copy always starts where it left off." + GetSummaryVar;
	else if (optUpdate.checked) {
		GetSummaryVar = "\nExisting files will be updated. No new files will be added." + GetSummaryVar;
		GetSummaryVar = "Files will be updated based on their 'modified' dates." + GetSummaryVar;
	}else if (optClean.checked){
		GetSummaryVar = "of target directories will be deleted if they exist." + GetSummaryVar;
		GetSummaryVar = "No overwriting or comparison will take place. Note that all items inside " + GetSummaryVar;
		GetSummaryVar = "Destination files and directories will be deleted before copy occurs. " + GetSummaryVar;
	}
	//functionality duplicatged with the test button
	/*if (chkTest.checked) GetSummaryVar = "----TRIAL RUN----\n" + GetSummaryVar;*/
	return GetSummaryVar;
}

//Compiles a list of arguments for each xcopy command based on user preferences
function GetArgs(){
	var GetArgsVar = "";
	
	if (chkArchiveTag.checked)
		//Only items tagged 'archive' will be copied.
		//The 'archive' attribute will be copied to the destination files.
		//The 'archive' attribute will be turned off on the source files.
		GetArgsVar += " /M";
	else if (chkArchiveOnly.checked)
		//Only items tagged 'archive' will be copied.
		//The 'archive' attribute will be copied to the destination files.
		GetArgsVar += " /A";
	if (chkDecrypt.checked)
		//Items will be decrypted during the transfer.
		GetArgsVar += " /G";
	if (chkHiddenSystem.checked)
		//Hidden files and System files will be copied.
		GetArgsVar += " /H";
	if (chkResume.checked)
		//Copy will resume after a connection failure.
		GetArgsVar += " /Z";
	if (chkMaintainReadOnly.checked)
		//The read-only attribute for files will be maintained.
		GetArgsVar += " /K";

	if (chkMaintainAudit.checked)
		//All file audit and security information will be maintained.
		GetArgsVar += " /X";
	else if (chkMaintainSecurity.checked)
		//Security and file ownership information will be maintained.
		GetArgsVar += " /O";
	
	
	//Errors will be ignored.
	GetArgsVar += " /C";

	//file contents piping not allowed.
	GetArgsVar += " /I";
	
	if (optOverwrite.checked) {
		//Source files will overwrite destination files.
		GetArgsVar += " /R /Y" ;
	}else if (optContinue.checked){
		//If destination files are older, they will be overwritten.
		GetArgsVar += " /R /D /Y";
	}else if( optContinue2.checked){
		//If destination files are older, they will be overwritten.
		GetArgsVar += " /R /D /Y" ;
	}else if (optUpdate.checked) {
		//Existing files will be updated. No new files will be added.
		//Files will be updated based upon a comparison of their 'modified' dates.
		GetArgsVar += " /R /U /Y" ;
	}else if (optClean.checked){
		//This will remove files in subdirectories in the source file set that have been deleted from the destination.
		//Destination files and directories will be deleted before copy occurs.
		GetArgsVar += " /R /Y" ;
	}
	
	return GetArgsVar;
}

//These are to convert filenames so that HTML doesn't choke on apostrophes and slashes
function EncodeString(value){
	var newvalue = value;
	var i;
	for (i=0;i<newvalue.length;i++){
		if (newvalue.charAt(i)=="\\"){
			newvalue = newvalue.substr(0,i) + "<" + newvalue.substring(i + 1,newvalue.length);
		}else if (newvalue.charAt(i)=="'"){
			newvalue = newvalue.substr(0,i) + ">" + newvalue.substring(i + 1,newvalue.length);
		}
	}
	return newvalue;
}
function DecodeString(value){
	var newvalue = value;
	var i;
	for (i=0;i<newvalue.length;i++){
		if (newvalue.charAt(i)=="<"){
			newvalue = newvalue.substr(0,i) + "\\" + newvalue.substring(i + 1,newvalue.length);
		}else		if (newvalue.charAt(i)==">"){
			newvalue = newvalue.substr(0,i) + "'" + newvalue.substring(i + 1,newvalue.length);
		}
		
	}
	return newvalue;
}
