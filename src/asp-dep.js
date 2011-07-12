var version='0.5';
var LOG_NONE=-1, LOG_ERROR=0, LOG_WARN=1, LOG_INFO=2;
var logLevel=LOG_WARN;
/* History:
 * v0.5: Added <script runat=server> parsing 
 * v0.4: Fixed a bug with -run
 *       Fixed a bug with -absolute (relative paths were not made absolute)
 *       Added -dups and -nodups, -indent, -noindent
 *       Renamed options to short syntax with number (-relative is now -r1 and -absolute is -r0, etc.)
 *       Fixed a bug with -r1. NormalizePath ignored paths with a parent-path selector at the end (c:\temp\..)
 *       Added -f, to lookup definition of function/sub
 *       Changed default value of count and dups
 *       Fixed bug (when the file-attribute of the #include directive wasn't specified in lowercase, the file didn't exist)
 *       Fixed bug (whitespace before/after the = sign after file/virtual
 * v0.3: added functions: -relative, -absolute, -start, -run
 *       Added bare logging.
 * v0.2: Made virtual-directory optional
 *       Now also recognizes filenames in include directives with no quotes (ie: <!--#include virtual=/include/nocache.inc--> )
 *       Fixed a bug with forwardslashes in file-includes.
 * v0.1: Initial
 *
 * Bugs/feature requests:
 * -SOLVED: Doesn't parse <script runat=server
 * -SOLVED: resolve where a function-name is defined: vbscript (-f option)
 * -resolve where a function-name is defined for javascript
*/
var cdir, vdir, incs;
var indentCount=0, indentDelta=2;
var showLogo=true;
var relative=true;
var count=false;
var dups=false;
var icon=true;
var indent=true;
var definition, defs=[]; //lookup function or sub. Definition is string to match sub or function, defs is list of found files
var start;

Main(WScript.Arguments);

function Main(argv) {
	ShowLogo();
	cdir=CurrentDir(); 
	//Handle arguments
	if(argv.length<1) ShowHelp('You should supply at least one argument');
	var argn=0;
	var file=argv(argn++);
	if(!FileExists(file)) {
		error('Supplied file does not exist: '+vdir);
		ShowHelp();
	}
	if(argv.length>1&&argv(argn).substr(0,1)!='-') {
		vdir=NormalizePath(argv(argn++));
		if(!DirExists(vdir)) {
			error('Supplied virtual directory does not exist: '+vdir);
			ShowHelp();
		}
	}
	else {
		vdir=cdir;
	}
	if(vdir.substr(vdir.length-1,1)!="\\") vdir+="\\";

	for(var i=argn; i<argv.length; i++) {
		var arg=argv(i);
		switch(arg) {
		case '-b0': case '-b1':
			icon=arg.substr(2,1)=='1'; 
			info('set '+(icon?'icon':'noicon')+' (bullet)');
			break;
		case '-c0': case '-c1':
			count=arg.substr(2,1)=='1'; 
			info('set '+(count?'count':'nocount'));
			break;
		case '-d0': case '-d1':
			dups=arg.substr(2,1)=='1'; 
			info('set '+(dups?'dups':'nodups'));
			break;
		case '-f':
			if(i+1>argv.length) {
				ShowHelp('-f specified, but no name');
			}
			else {
				definition=argv(++i);
				info('Searching for definition '+definition);
			}
		case '-i0': case '-i1':
			indent=arg.substr(2,1)=='1'; 
			info('set '+(indent?'indent':'noindent'));
			break;
		case '-l0': case '-l1':
			showLogo=arg.substr(2,1)=='1'; 
			info('set '+(indent?'showlogo':'nologo'));
			break;
		case '-nologo':
			showLogo=false;
  		info('set nologo');
			break;
		case '-r0': case '-r1':
			relative=arg.substr(2,1)=='1'; 
			info('set '+(relative?'relative':'absolute'));
			break;
		case '-v0':	case '-v1':	case '-v2':
			logLevel=+arg.substr(2,1);
		  info('set logLevel '+['error','warn','info'][logLevel]);
			break;
		case '-start':
		case '-run':
			if(i+1>=argv.length) {
				var s=GetEnv('EDITOR');
				if(s) {
					info('set start from environment');
					start=s;
				}
				else {
					ShowHelp('-start parameter specified, but not a program_name (and no environment variable EDITOR specified)');
				}
			}
			else {
				info('set start from arguments');
				start=argv(++i);
			}
			break;
		default:
			warn('Unrecognized argument: '+arg);
			break;
		}
	}
	incs=new Object(); //hash array with all includes
	ParseDep(cdir+file,vdir);
	if(definition) {
		if(defs.length==0) {
			WScript.Echo('Definition ['+definition+'] not found.');
		}
		else {
			WScript.Echo('Definition ['+definition+'] found in '+defs.length+' file'+(defs.length==1?'':'s')+':');
			for(var i=0; i<defs.length; i++) {
				WScript.Echo("  "+defs[i]);
			}
		}
	}
	if(start) {
		for(var i in incs) {
			var runThis='"'+start+'" '+i;
			info("RUN: "+runThis);
			Execute(runThis);
		}
	}
}

function ShowLogo() {
	WScript.Echo("asp-dep -- versie "+version+" -- Show ASP Include dependencies -- By Catsdeep, 2005");
	WScript.Echo("");
}

function ShowHelp(s) {
	if(s) warn(s+'\n');
	WScript.Echo("Usage: asp-dep.js asp-file.asp [virtual-directory] {-(b|c|d|i|r)(0|1)} [-v(0|1|2)] [-f function_name] [-run [program_name]]");
	WScript.Echo("");
	WScript.Echo("When the virtual-directory is not supplied, the current directory is assumed to be the virtual directory");
	WScript.Echo("-b(0|1) Show bullets (1, default) or don't (0)");
	WScript.Echo("-c(0|1) Count and display (1) the number of includes or don't (0, default)");
	WScript.Echo("-d(0|1) Show duplicate include entries (1) or don't (0, default)");
	WScript.Echo("-i(0|1) Indent include-level (1, default) or don't (0)");
	WScript.Echo("-r(0|1) Paths are shown relative (0, default) to the virtual dir, when possible");
	WScript.Echo("        or absolute (1) (full path, including drive name)");
	WScript.Echo("-v(0|1|2) Set verbosity level. 0:Errors, 1:warnings (default), 2: info");
	WScript.Echo("-f func Lookup file of definition of the specified function (VB style)");
	WScript.Echo("-run    Run the program 'program_name' with all found files as parameter");
	WScript.Echo("");
	WScript.Echo("In combination with -start, you can also specify a program_name with the\nenvironment-variable EDITOR.");
	WScript.Echo("");
	WScript.Echo("Example.");
	WScript.Echo("This commands opens the asp-file index.asp and all it's includes with notepad:");
	WScript.Echo("  asp-dep.js index.asp -start notepad.exe");
	WScript.Quit(1);
}

function MinusVDir(s) {
	s=''+s;
	if(relative) {
		if(s.toLowerCase().indexOf(vdir.toLowerCase())==0) return (''+s).substr(vdir.length);
		else return s; //not in virtual directory?
	}
	else {
		return s; //absolute
	}
}
function ParseDep(file) {
	function ShowFile(iconchar,remark) {
		var a=[];
		if(icon) a[a.length]=iconchar;
		a[a.length]=MinusVDir(file);
		if(remark!=null) a[a.length]='('+remark+')';
		WriteLine(a.join(' '));
	}
	file=NormalizePath(''+file);
	if(FileExists(file)) {
		if(incs[file.toLowerCase()]) {
			if(dups) ShowFile("*","duplicate, NO FOLLOW");
		}
		else {
			incs[file.toLowerCase()]=true; 
			var includes=GetIncludes(file);
			ShowFile((includes.length==0?"-":"+"),count?includes.length+" includes":null);
			indentCount+=indentDelta;
			for(var i=0; i<includes.length; i++) {
				ParseDep(includes[i]);
			}
			indentCount-=indentDelta;
		}
	}
	else {
		ShowFile("#","DOES NOT EXIST");
	}
}
function GetPath(s) {
	return (''+s).substring(0,s.lastIndexOf("\\")+1);
}
function NormalizePath(s) {
	if(s.substr(1,1)!=':') {
		s=cdir+'\\'+s;
	}
	s=s.replace(/\\\\/g,"\\"); //remove double backslashes
	s=s.replace(/\//g,'\\'); //convert forwardslashes to backslashes
	while(s.indexOf('\\..\\')!=-1) {
		s=s.replace(/[^\\]+\\\.\.\\/gi,''); //normalize parent paths
	}
	s=s.replace(/[^\\]+\\\.\.$/,''); //als .. het laatste stuk van het pad is, verwijderen.
	return s;
}
function /*constructor*/ IncludeEntry(type,name,infile) {
	this.type=(''+type).toLowerCase();
	this.name=name.replace(/\//g,'\\');
	this.toString=function() {
		if(this.type=='file') {
			return GetPath(NormalizePath(infile))+this.name;
		}
		else return vdir+this.name;
	}
}
function GetIncludes(file) {
  function GetAttr(line,attr) {
    var re=new RegExp(attr+"\\s*=\\s*('\\s*([^']+)\\s*'|\"\\s*([^\"]+)\\s*\"|\\s*([^\\s>]+))", "i");
    var a=re.exec(line);
    if(a) return (a[2]||a[3]||a[4]);
    else return null;
  }
	var txt=ReadFile(file);
	//--| Handle includes (ssi syntax)
	var re=/<!--\s*#include\s+(virtual|file)\s*=\s*("([^"]+)"|[^ ]+)\s*-->/gi;
	var result=[], a;
	while(a=re.exec(txt)) {
		result[result.length]=new IncludeEntry(a[1],a[3]?a[3]:a[2],file);
	}
	//--| Handle includes (<script runat=server syntax)
  re=/<script[^>]+/gi;
  while(a=re.exec(txt)) {
    var line=a[0];
    var runat=GetAttr(line,"runat");
    if(runat&&runat.toLowerCase()=='server') {
      var src=GetAttr(line,"src");
      if(src) {
        //ignore the language
        result[result.length]=new IncludeEntry("virtual",src,file);
      }
    }
  }
	//--| Get definitions, VB style
	if(definition) {
		var reVbDef=/(Function|Sub)[ \t]+([a-z][a-z0-9_]*)/gi;
		info('Parsing definitions for file '+file);
		while(a=reVbDef.exec(txt)) {
			info('Found definition: '+a[2]);
			if(a[2].toLowerCase()==definition.toLowerCase()) {
				defs[defs.length]=NormalizePath(''+file);
			}
		}
	}
	return result;
}

function WriteLine(s) {
	WScript.Echo(Indent(indentCount)+s);
}
function Indent(n) {
	return indent?new Array(n+1).join(' '):'';
}
function FileExists(filespec) {
   return new ActiveXObject("Scripting.FileSystemObject").FileExists(filespec)
}
function DirExists(filespec) {
   return new ActiveXObject("Scripting.FileSystemObject").FolderExists(filespec)
}
function ReadFile(filename) {
   var fso=new ActiveXObject("Scripting.FileSystemObject");
   var f=fso.OpenTextFile(filename, 1);
	 var s=f.ReadAll();
	 f.Close();
	 return s;
}
function CurrentDir() {
	cdir=WScript.CreateObject ("WScript.Shell").CurrentDirectory; 
	if(cdir.substr(cdir.length-1,1)!="\\") cdir+="\\";
	cdir=NormalizePath(cdir);
	return cdir;
}
function Execute(s) {
	var shell = new ActiveXObject("WScript.Shell");
	//shell.Exec(s);
	shell.Run(s,8,true);
}
function GetEnv(s) {
	var shell=WScript.CreateObject("WScript.Shell");
	var env=shell.Environment("USER");
	return env(s);
}
function info(s)  { if(logLevel>=LOG_INFO)  WScript.Echo('INFO: '+s); }
function warn(s)  { if(logLevel>=LOG_WARN)  WScript.Echo('WARN: '+s); }
function error(s) { if(logLevel>=LOG_ERROR) WScript.Echo('ERROR: '+s);  }