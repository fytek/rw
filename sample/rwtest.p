DEFINE VARIABLE RWhandle AS COM-HANDLE.
DEFINE VARIABLE RWres AS COM-HANDLE.

CREATE "FyTek.ReportWriter" RWhandle.

RWhandle:setOutFile ("c:\temp\test.pdf").
RWhandle:setExe("c:\<path to program>\pdfrw64.exe").

/* 
Normally you would start the server from another process.
Just showing an examlpe how you might do it from Progress.
Note you don't have to use a server but if you are sending multiple
requests it will prevent resource overload.
*/
RWhandle:setKeyName("demo"). /* Needed to test the server start */
RWhandle:startServer ("localhost",7075,5,"c:\temp\logs.txt").

 /* 
 This is necessary when the server is started elsewhere.
 Just tells the program what host and port the server is listening on.
*/
RWhandle:setServer("localhost",7075).

RWhandle:setComp15().

RWhandle:setPDFCmd ("<PDF>").
RWhandle:setPDFCmd ("<PAGE>").
RWhandle:setPDFCmd ("<TEXT ALIGN=C>").
RWhandle:setPDFCmd ("Hello, world").
RWhandle:setPDFCmd ("</TEXT>").

/* Tells the DLL to build the PDF */
RWres = RWhandle:buildReport().
message RWres:getMsg().
/* RWres:getBytes() has the bytes of the PDF if true is passed as a parameter to buildReport. */

/* Normally you would not do this here.  
The server would remain running for other clients. */
RWhandle:stopServer().

RELEASE OBJECT RWhandle.

