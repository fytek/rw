
Dim rwObj, cmd, srvOpts

' Make sure to register the dll first, if necessary (your path might be different)
' Microsoft.Net.Compilers.3.4.0\tools\csc /target:library /platform:anycpu /out:pdfrw_20.dll pdfrw_20.cs /keyfile:mykey.snk
' C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase c:pdfrw_20.dll
' cscript.exe (or wscript.exe) rwtest.vbs
set rwObj = CreateObject("FyTek.ReportWriter")

' Running as a server is optional. You may just call the program directly if you prefer.
' Running as a server provides control over resources and single copy of the program exists
' in memory.  The server may be on any box that is accsible from the one making the request.
' 
' In this sample we are starting a server to show how it is done and then shutting it down 
' with this script finishes.  In a production environment you would likely have these in 
' separate processes - that is, have one process that only starts or stops the server.  Other
' processes access the server to build PDFs and don't shut it down when they are done.  It 
' stays running for other processes to come along and access it.

' Set the location of the executable - only needed if we are starting the server or not using 
' a server.  When a server is used, only the host and port are needed.
' rwObj.setExe(".\pdfrw64")
 
' Needed to test server mode - in produciton, use setKeyName and setKeyCode with your values.  
' Or use licInfo method to set your license info.
' Assumes we're using the demo, in a real situation, include setKeyCode also.
' And if using an already running server, no key name is necessary.
rwObj.setKeyName("demo") 

' Start the server.  Normally you would do this elsewhere, in another program
' or from a shell script, etc.  You might have several servers running for load balancing.
srvOpts = rwObj.startServer(,,,"c:\temp\mytestlog.txt") 

' The setServerFile method is used to pass in a file name of IP addresses and ports that
' you have one or more servers running on.  This is static for the DLL so all users accessing
' the DLL will have the same file.  The DLL will then pass requests to different servers
' depending on load.
' For example, in serverfile.dat you might have:
' # here are the RW servers
' 192.168.1.124 7075
' 192.168.1.125 7075
' localhost 7075
' rwObj.setServerFile("servers.dat") 

' In a real situation where server is already running and there is no server
' setting file (setServerFile call that provides a file of one or more Report Writer servers)
' then you might need to set the values here so the program knows the box and port to use.
' For example, when the server is running off of a Linux server on box 192.168.1.124:
' rwObj.setServer "192.168.1.124",7075
' If the Report Writer server is on a different box (not localhost for example) then you might
' want to set this option that will auto send any files from this box to the server so it has
' all the needed files to build the PDF.  Or, use the sendFileTCP method to provide the files
' one at a time.  If more than one or you have included images, fonts, then may be easier to juse
' use setAutoSendFile.
rwObj.setAutoSendFiles()
cmd = rwObj.resetOpts(true)

WScript.Echo("Building PDF1 - assumes there is a file c:\temp\sample.frw")
' Now that server is running, send some samples
cmd = rwObj.setInFile("abc.frw") ' This file does't exist
cmd = rwObj.sendFileTCP("abc.frw","c:\temp\sample.frw") ' Here is the file to pass as abc.frw
' We could have passed c:/temp/sample.frw in setInFile if the server is running on the same
' box that we're running this script from.  You only need sendFileTCP if the Report Writer server
' is running on a different box so it knows to copy the file over.  Otherwise, the path and file
' will be assumed to be on the box where Report Writer server is running.

' Create the PDF and get back the byte array - or pass false to not wait for the PDF to be returned
Set d = rwObj.buildPDFTCP (true)
' d.getByptes() will contain the newly build PDF - we'll save to a file but you might want to 
' display on a web page or save in a database
if (StrComp(d.Msg,"OK") = 0) Then
    SaveBinaryDataTextStream "c:\temp\myfile.pdf", d.Bytes
End If

' clear out the commands for another run        
cmd = rwObj.resetOpts()
' let's check the server status (if multiple servers, pass true to get info on all of them)
cmd = rwObj.serverStatus(true)
WScript.Echo(cmd)

cmd = rwObj.setPDFCmd("<PDF>")
cmd = rwObj.setPDFCmd("<PAGE>")
cmd = rwObj.setPDFCmd("<TABLE>")
cmd = rwObj.setPDFCmd("<TR>")
cmd = rwObj.setPDFCmd("<TD>")
cmd = rwObj.setPDFCmd("&date &time")
cmd = rwObj.setPDFCmd("</TD>")
cmd = rwObj.setPDFCmd("<TD>")
cmd = rwObj.setPDFCmd("This is a test string")
cmd = rwObj.setPDFCmd("</TD>")
cmd = rwObj.setPDFCmd("</TR>")
cmd = rwObj.setPDFCmd("</TABLE>")

WScript.Echo("Building PDF2...")
Set d = rwObj.buildPDFTCP(true) 
cmd = rwObj.resetOpts()

' Saving the PDF in this program
if (StrComp(d.Msg,"OK") = 0) Then
    SaveBinaryDataTextStream "c:\temp\myfile2.pdf", d.Bytes
End If    

cmd = rwObj.setPDFCmd("<PDF>")
cmd = rwObj.setPDFCmd("<PAGE>")
cmd = rwObj.setPDFCmd("<TEXT>")
cmd = rwObj.setPDFCmd("This one is saved from the DLL")
cmd = rwObj.setPDFCmd("</TEXT>")

WScript.Echo("Building PDF3...")

' Filepath and name are supplied so it will be saved locally here
Set d = rwObj.buildPDFTCP(false, "c:\temp\myfile3.pdf") 
cmd = rwObj.resetOpts()

' shut down the server - typically you would leave it running for the next process to access however
srvOpts = rwObj.stopServer() 

' Helper functions for saving the file to disk
Function RSBinaryToString(xBinary)
  Dim Binary
  'MultiByte data must be converted To VT_UI1 | VT_ARRAY first.
  If vartype(xBinary)=8 Then Binary = MultiByteToBinary(xBinary) Else Binary = xBinary
  
  Dim RS, LBinary
  Const adLongVarChar = 201
  Set RS = CreateObject("ADODB.Recordset")
  LBinary = LenB(Binary)
  
  If LBinary>0 Then
    RS.Fields.Append "mBinary", adLongVarChar, LBinary
    RS.Open
    RS.AddNew
      RS("mBinary").AppendChunk Binary 
    RS.Update
    RSBinaryToString = RS("mBinary")
  Else
    RSBinaryToString = ""
  End If
End Function

Function SaveBinaryDataTextStream(FileName, ByteArray)
  ' This function is to write a PDF to disk but
  ' you could instead send to a web page
  ' or store in a database, etc.

  'Create FileSystemObject object
  Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
  
    If FS.FileExists(FileName) Then
      FS.DeleteFile FileName
    End If

  'Create text stream object
  Dim TextStream
  Set TextStream = FS.CreateTextFile(FileName,ForWriting)
  
  'Convert binary data To text And write them To the file
  TextStream.Write (RSBinaryToString (ByteArray))
End Function
