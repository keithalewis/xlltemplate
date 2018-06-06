# xlltemplate

Template for Excel add-ins

Fork or clone the repository. Set the `Configuration` (Debug or Release) and `Platform` (x86 or x64) and build the `template` project.  
The platform must be the same as Excel &endash; x86 for 32-bit or x64 for 64-bit.

## Debugging

To debug an add-in you must tell Visual Studio the full path to the Excel executable and what add-in to load.  
In the project `Properties` in the `Debugging` tab the `Command` should be the full path to the Excel executable.  
You can `<Browse...>` to it using the dropdown or obtain it from the Task Manager if Excel is running. Under `Processes`
right click on `EXCEL.EXE` and choose `Open File Location`.  
To have Excel open your add in when debugging specify `Command Arguments` to be `"$(TargetPath)"`.  
You may want to also specify `/p $(ProjectDir)` so Excel will default to opening files in your project directory.
