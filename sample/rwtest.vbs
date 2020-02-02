
Dim mObj, cmd, pdf, srvOpts

' Make sure to register the dll first, if necessary (your path might be different)
' Microsoft.Net.Compilers.3.4.0\tools\csc /target:library /platform:anycpu /out:pdfrw_20.dll pdfrw_20.cs /keyfile:mykey.snk
' C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase c:pdfrw_20.dll
' cscript.exe (or wscript.exe) rwtest.vbs
set mObj = CreateObject("FyTek.ReportWriter")

' Needed to test server mode - in produciton, use setKeyName and setKeyCode with your values.  
' Or use licInfo method to set your license info.
' Running as a server is optional. You may just call the program directly if you prefer.
' Running as a server provides control over resources and single copy of the program exists
' in memory.
' In this sample we are starting a server to show how it is done and then shutting it down 
' with this script finishes.  In a production environment you would likely have these in 
' separate processes - that is, have one process that only starts or stops the server.  Other
' processes access the server to build PDFs and don't shut it down when they are done.  It 
' stays running for other processes to come along and access it.
mObj.setKeyName("demo") ' Assumes we're using the demo, in a real situation, include setKeyCode also.
srvOpts = mObj.startServer(,,,"c:\temp\mytestlog.txt") 

' Now that server is running, send some samples
cmd = mObj.setInFile("abc.frw") ' This file does't exist
cmd = mObj.sendFileTCP("abc.frw","c:/temp/sample.frw") ' Here is the file to pass as abc.frw
' We could have passed c:/temp/sample.frw in setInFile if the server is running on the same
' box that we're running this script from.  You only need sendFileTCP if the Report Writer server
' is running on a different box so it knows to copy the file over.  Otherwise, the path and file
' will be assumed to be on the box where Report Writer server is running.

' Create the PDF and get back the byte array - or pass false to not wait for the PDF to be returned
pdf = mObj.buildPDFTCP(true) 
' pdf will contain the newly build PDF - we'll save to a file but you might want to 
' display on a web page or save in a database
SaveBinaryDataTextStream "c:\temp\myfile.pdf", pdf

' clear out the commands for another run        
cmd = mObj.resetOpts()
' let's check the server status
cmd = mObj.serverStats()

WScript.Echo(cmd)

' shut down the server - typically you would leave it running for the next process to access however
srvOpts = mObj.stopServer() 

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
