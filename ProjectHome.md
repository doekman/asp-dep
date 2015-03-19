When fixing code on legacy ASP code, you can easily loose track of all included files. With a given ASP-file, this utility can find all referenced files.

## Details ##

Both VBScript and JavaScript as target language are supported. The tool can be used for:

  * Find missing and duplicate includes
  * Find the file(s), in which a function is defined (VBScript only)
  * Process both `virtual` as `file` includes (server side include syntax)
  * Also support `<script runat=server>` syntax
  * Combines all dependencies to one file

## Command-line options ##

```
Usage: asp-dep.js asp-file.asp [virtual-directory] {-(b|c|d|i|r)(0|1)} [-v(0|1|2)] [-f function_name] [-run [program_name]]

When the virtual-directory is not supplied, the current directory is assumed to be the virtual directory
-b(0|1) Show bullets (1, default) or don't (0)
-c(0|1) Count and display (1) the number of includes or don't (0, default)
-d(0|1) Show duplicate include entries (1) or don't (0, default)
-i(0|1) Indent include-level (1, default) or don't (0)
-r(0|1) Paths are shown relative (0, default) to the virtual dir, when possible
        or absolute (1) (full path, including drive name)
-v(0|1|2) Set verbosity level. 0:Errors, 1:warnings (default), 2: info
-f func Lookup file of definition of the specified function (VB style)
-run    Run the program 'program_name' with all found files as parameter

In combination with -start, you can also specify a program_name with the
environment-variable EDITOR.

Example.
This commands opens the asp-file index.asp and all it's includes with notepad:
  asp-dep.js index.asp -start notepad.exe
```

## Final note ##

This script (written in JavaScript with WScript as executing environment) is written some years ago, and provided to be very helpful at the time. However, the code is not production quality. Also, the script is no longer maintained.