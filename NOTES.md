Use MSBuild to create documentation.

%ProgramFiles(x86)%\MSBuild\14.0\bin\MSBuild.exe
    /p:CopyrightText="Copyright \xa9 2006-2011"
    /p:FeedbackEMailAdress="Eric@EWoodruff.us"
    HelpFile.shfbproj