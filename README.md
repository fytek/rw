# rw
FyTek PDF Report Writer DLL

This program is supplied as a compiled .NET DLL assembly with the executable download of Report Writer.  You may download this source and compile yourself (instructions are at the top of the code).

The purpose is to allow a DLL method interface to the executable program pdfrw.exe or pdfrw64.exe.  This DLL, which is compiled as both a 32-bit and 64-bit version, may be used in Visual Basic program, ASP, C#, etc. to send the information for building and retreiving PDFs from PDF Report Writer.  This DLL also replaces the old DLL version that was limited to running Report Writer on the same box and had occassional memory issues.

There are commands to start a Report Writer server which loads the executable version of Report Writer in memory and listens on the port you specify for commands.  This also allows you to have Report Writer running on a different server for load balancing by not building the PDF on the same box as the requestor.  You may also run several instances of Report Writer server all on different boxes if you wish and the DLL will cycle between them to handle requests.
