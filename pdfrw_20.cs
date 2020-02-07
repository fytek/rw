// pdfrw_20.cs - .NET DLL wrapper for FyTek's PDF Report Writer

// Call this DLL from the language of your choice as long as it supports
// a COM or .NET DLL.  May be 32 or 64-bit.
// This DLL accepts parameters and then builds the PDF which you may save
// locally, on the box the server is running on (if it's different box) or
// have the PDF returned as a byte array for display on website or for saving
// in a database.
// This DLL calls the Report Writer executable (32 or 64-bit) or uses sockets
// when Report Writer is running as a server.  It sends the parameter settings
// made here to build the PDF.  For exapmle, you might want to startup Report Writer
// with a pool of 5 connections on a Linux box and call it from this DLL on a
// Windows box.  Even if you run Report Writer on the same box it's recommended
// to start a Report Writer server in order to keep resource usage in check.
// Note you may start more than one Report Writer server at a time with different
// port numbers for each one.
// Use startServer to start up a Report Writer server and stopServer to shut it down.
// You probably want to do that outside of your main routine that will be building
// PDFs as your main routine will link to this DLL to call the already running
// service.


using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Security.Cryptography;

// Compiling this code:
// Microsoft.Net.Compilers.3.4.0\tools\csc /target:library /platform:anycpu /out:pdfrw_20.dll pdfrw_20.cs /keyfile:mykey.snk
// C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase c:pdfrw_20.dll
// cscript.exe (or wscript.exe) rwtest.vbs

namespace FyTek
{
    [ComVisible(true)]
    [Guid("B6B57ADA-2D6A-491E-8B60-E8C3BDEAE4C4")]
	[ProgId("FyTek.ReportWriter")]    
    public class ReportWriterPDF : IDisposable
    {
        void IDisposable.Dispose()
        {

        }
        private TcpClient client = new TcpClient();
        private NetworkStream stream;
        private List<string> cmds = new List<string>(); // input commands (input.frw)
        private List<string> dataCmds = new List<string>(); // data file (when not using setDataFile)
        private String exe = "pdfrw64"; // the executable - change with setExe
        private Dictionary<string,string> opts = new Dictionary<string,string>(); // all of the parameter settings from the method calls
        private Dictionary<string,object> server = new Dictionary<string,object>(); // the server host/port/log file key/values
        
        // Start up Report Writer as a server
        [ComVisible(true)]
        public String startServer(
            String host = "localhost",
            int port = 7075,
            int pool = 5,
            String log = ""
          )
       {
            byte[] bytes = {};
            String errMsg = "";

            server.Clear();
            server.Add("host",host);
            server.Add("port",port);
            server.Add("pool",pool);
            server.Add("log",log); // file on server to log the output
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = true;
            startInfo.FileName = exe;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            String cmdsOut = "";
            String s = "";

            if (opts.TryGetValue("licInfo_name", out s)){
                String p = "";
                String d = "";
                opts.TryGetValue("licInfo_name", out p);
                opts.TryGetValue("licInfo_autodl", out d);
                cmdsOut += " -licname " + s + " -licpwd " + p + (d.Equals("1") ? " -licweb" : "");
            }
            if (opts.TryGetValue("keyCode", out s))
                cmdsOut += " -kc " + s;
            if (opts.TryGetValue("keyName", out s))
                cmdsOut += " -kn " + s;

            cmdsOut += (!log.Equals("") ? " -log " + '"' + log + '"' : "")
                + " -port " + port + " -pool " + pool + " -host " + host;
            startInfo.Arguments = "-server " + cmdsOut;
            (bytes, errMsg) = runProcess(startInfo, false);
            opts.Remove("licInfo_name");
            opts.Remove("licInfo_pwd");
            opts.Remove("licInfo_autodl");
            opts.Remove("keyName");
            opts.Remove("keyCode");
            return errMsg;
        }

        [ComVisible(true)]
        public void setServer(
            String host = "localhost",
            int port = 7075
          )
       {
            server["host"] = host;
            server["port"] = port;
       }

        // Stop the server
        [ComVisible(true)]
        public String stopServer(){
            byte[] bytes;
            String errMsg;
            setOpt("serverCmd","-quit");
            (bytes, errMsg) = callTCP(isStop: true);
            return errMsg;
        }

        // Get the current stats for the server
        [ComVisible(true)]
        public String serverStatus(){
            byte[] bytes;
            String errMsg;
            setOpt("serverCmd","-serverstat");
            (bytes, errMsg) = callTCP(isStatus: true);
            return errMsg;
        }

        // Stop a process
        [ComVisible(true)]
        public String serverCancelId(int id){
            byte[] bytes;
            String errMsg;
            setOpt("stopId","-stopid " + id);
            (bytes, errMsg) = callTCP();
            return errMsg;
        }

        // Build the PDF using the server, optionally return the 
        // raw bytes of the PDF for further processing when retBytes = true
        // optional file name as well if saveFile is passed - this allows
        // for saving file on this box if server is running on different box
        [ComVisible(true)]
        public byte[] buildPDFTCP(bool retBytes, String saveFile = ""){            
            byte[] bytes = {};
            String errMsg = "";
            String s = "";

            if (cmds.Count > 0){
                if (!opts.TryGetValue("inFile", out s)){
                    s = $@"{Guid.NewGuid()}.frw"; // come up with unique file name
                    setInFile(s);
                    sendFileTCP(s, "--input--commands--");
                }
            }

            if (dataCmds.Count > 0){
                if (!opts.TryGetValue("dataFile", out s)){
                  s = $@"{Guid.NewGuid()}"; // come up with unique file name
                  setCmdlineOpts("-data " + s);
                  sendFileTCP(s, "--data--commands--");
                }
            }
            if (!saveFile.Equals("") || retBytes){
                setOutFile("genfile");
            }

            (bytes, errMsg) = callTCP(retBytes: retBytes, saveFile: saveFile);
            return bytes;
        }

        // Send all files to server - only necessary if server is on a different box
        [ComVisible(true)]
        public void setAutoSendFiles(){
          setOpt("autoSend","Y");
        }

        // Pass file contents in memory for input type files
        [ComVisible(true)]
        public String setDataCmd(String a)
        {
            dataCmds.Add(a);
            return a;
        }

        // Pass data file name
        [ComVisible(true)]
        public String setDataFile(String a)
        {
            setOpt("dataFile",a);
            return a;
        }

        // Pass file contents in memory for input type files
        [ComVisible(true)]
        public String setPDFCmd(String a)
        {
            cmds.Add(a);
            return a;
        }

        // Assign the key name
        [ComVisible(true)]
        public String setKeyName(String a)
        {
            if (a.ToLower().Equals("demo")){
                // Get the demo key from website - this only works with the demo pdfrw executable
                WebClient client = new WebClient();
                string res = client.DownloadString("http://www.fytek.com/cgi-bin/genkeyw_v2.cgi?prod=reportwriter");
                Regex regex = new Regex("-kc [A-Z0-9]*");
                Match match = regex.Match(res);                
                setOpt("keyName","testkey");
                setOpt("keyCode",match.Value.Substring(4));
            } else {
                setOpt("keyName",a);
            }
            return a;
        }

        // Assign the key code
        [ComVisible(true)]
        public String setKeyCode(String a)
        {
            setOpt("keyCode",a);
            return a;
        }

        // License
        [ComVisible(true)]
        public void licInfo(String licName,
            String licPwd,
            int autoDownload)
        {
            setOpt("licInfo_name",licName);
            setOpt("licInfo_pwd",licPwd);
            setOpt("licInfo_autodl",$"{autoDownload}");
        }

        // Assign the user
        [ComVisible(true)]
        public String setUser(String a)
        {
            setOpt("user",a);
            return a;
        }

        // Assign the owner
        [ComVisible(true)]
        public String setOwner(String a)
        {
            setOpt("owner",a);
            return a;
        }

        // Set no annote
        [ComVisible(true)]
        public void setNoAnnote()
        {
            setOpt("noAnnote","Y");
        }

        // Set no copy
        [ComVisible(true)]
        public void setNoCopy()
        {
            setOpt("noCopy","Y");
        }

        // Set no change
        [ComVisible(true)]
        public void setNoChange()
        {
            setOpt("noChange","Y");
        }

        // Set no print
        [ComVisible(true)]
        public void setNoPrint()
        {
            setOpt("noPrint","Y");
        }

        // Set the GUI process window off
        [ComVisible(true)]
        public void setGUIOff()
        {
            setOpt("guiOff","Y");
        }

        // Set any other command line type options
        [ComVisible(true)]
        public String setCmdlineOpts(String a)
        {
            String s = "";
            if (opts.TryGetValue("extOpts", out s))
                a = s + " " + a;
            setOpt("extOpts",a);
            return a;
        }

        // Set the quick mode
        [ComVisible(true)]
        public void setQuick()
        {
            setOpt("quick","Y");
        }

        // Set the quick2 mode
        [ComVisible(true)]
        public void setQuick2()
        {
            setOpt("quick2","Y");
        }

        // Assign the executable
        [ComVisible(true)]
        public String setExe(String a)
        {
            exe = a;
            return a;
        }

        // Assign the input file name
        [ComVisible(true)]
        public String setInFile(String a)
        {
            setOpt("inFile",a);
            return a;
        }

        // Assign the output file name
        [ComVisible(true)]
        public String setOutFile(String a)
        {
            setOpt("outFile",a);
            return a;
        }

        // Compression 1.5
        [ComVisible(true)]
        public void setComp15(String a)
        {
            setOpt("comp15","Y");
        }

        // Encrypt 128
        [ComVisible(true)]
        public void setEncrypt128(String a)
        {
            setOpt("enc128","Y");
        }

        // AES 128
        [ComVisible(true)]
        public String setEncryptAES(String a)
        {
            setOpt("aes",a);
            return a;
        }

        // allow breaks
        [ComVisible(true)]
        public void setAllowBreaks(String a)
        {
            setOpt("allowBreaks","Y");
        }

        // open output
        [ComVisible(true)]
        public void setOpen()
        {
            setOpt("open","Y");
        }        

        // print output
        [ComVisible(true)]
        public void setPrint()
        {
            setOpt("print","Y");
        }        

        // Calls buildPDF - this is for legacy purposes
        [ComVisible(true)]
        public byte[] buildReport(bool waitForExit = true)
        {    
            return buildPDF(waitForExit);
        }

        // Call the executable (non server mode)
        [ComVisible(true)]
        public byte[] buildPDF(bool waitForExit = true)
        {    
            byte[] bytes = {};
            String errMsg = "";
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = !waitForExit;
            if (waitForExit){
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardInput = true;
            }
            startInfo.FileName = exe;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.Arguments = setBaseOpts();  
            (bytes, errMsg) = runProcess(startInfo,waitForExit);
            return bytes;
        }

        // Reset the options
        [ComVisible(true)]
        public void resetOpts(){
            String s;
            cmds = new List<string>();
            dataCmds = new List<string>();
            opts.Clear();
            opts.TryGetValue("inFile", out s);
            
        }

        // Passes file to server over socket
        public String sendFileTCP(String fileName,
            String filePath = ""){
            String errMsg = "";
            Byte[] data;
            Object host = "";
            // String to store the response ASCII representation.
            String responseData = String.Empty;
            String message = "";

            if (!server.TryGetValue("host", out host)){
                return "Server not setup - try calling setServer or startServer first.";
            }

            try {         
                if (!client.Connected) {
                    client = new TcpClient((String) host, (int) server["port"]);
                    stream = client.GetStream();
                }
                
                // Send a file
                byte[] buffer = new byte[1024];
                int bytesRead = 0;

                message = " -send --binaryname--" + fileName + "--binarybegin--";
                data = System.Text.Encoding.UTF8.GetBytes(message);   
                stream.Write(data, 0, data.Length);                    
                if (filePath.Equals("--data--commands--")){                        
                    foreach (String value in dataCmds){
                        data = System.Text.Encoding.UTF8.GetBytes(value);
                        stream.Write(data, 0, data.Length);
                    }
                } else if (filePath.Equals("--input--commands--")){ 
                    foreach (String value in cmds){
                        data = System.Text.Encoding.UTF8.GetBytes(value);
                        stream.Write(data, 0, data.Length);
                    }
                } else if (!filePath.Equals("")){
                    BinaryReader br;
                    br = new BinaryReader(new FileStream(filePath, FileMode.Open));

                    while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                        stream.Write(buffer, 0, bytesRead);                    

                }
                message = "--binaryend-- ";
                data = System.Text.Encoding.UTF8.GetBytes(message);             
                stream.Write(data, 0, data.Length);                           

            } catch (SocketException e) {
                errMsg = e.Message;
            } catch (IOException e) {
                errMsg = e.Message;
            }
            return (errMsg);
        }

        private void setOpt(String k, String v){
            opts[k] = v;
        }

        private (byte[], String) runProcess(ProcessStartInfo startInfo, bool waitForExit) {
            // Start the process with the info we specified.
            // Call WaitForExit and then the using statement will close.  
            byte[] bytes = {};
            try {              
                using (Process exeProcess = Process.Start(startInfo))
                {
                    if (waitForExit){
                        StreamWriter inputCmds = exeProcess.StandardInput;
                        foreach (String value in cmds){
                            inputCmds.WriteLine(value);
                        }
                        inputCmds.Close();
                    }

                    MemoryStream memstream = new MemoryStream();
                    byte[] buffer = new byte[1024];
                    int bytesRead = 0;
                    if (waitForExit){
                        BinaryReader br = new BinaryReader(exeProcess.StandardOutput.BaseStream);
                        while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                            memstream.Write(buffer, 0, bytesRead);                    
                    }                        
                    if (waitForExit){
                        exeProcess.WaitForExit();
                        bytes = memstream.ToArray();
                    }
                }
            }
            catch (Exception e){
                return (bytes, e.Message);
            }
            return (bytes, "");

        }

        // Passes data to server over socket but does not finalize (that is, send BUILDPDF command)
        private String sendTCP(){
            String errMsg = "";
            Byte[] data;
            Object host = "";
            // String to store the response ASCII representation.
            String responseData = String.Empty;
            String message = "";

            try {         
                
                String s;
                if (opts.TryGetValue("serverCmd", out s)){
                    message = s;
                    opts.Remove("serverCmd");
                } else {
                    message = setBaseOpts();
                }
                // Send commands
                data = System.Text.Encoding.UTF8.GetBytes(message);             
                stream.Write(data, 0, data.Length);      

            } catch (SocketException e) {
                errMsg = e.Message;
            } catch (IOException e) {
                errMsg = e.Message;
            }
            return (errMsg);
        }


        private String setBaseOpts(){
            String s = "";
            String message = "";
            if (opts.TryGetValue("inFile", out s))
                if (!s.Equals(""))
                    message += "\"" + s + "\" ";
            else if (cmds.Count > 0)        
                    message += "- "; // get from stdin
            if (opts.TryGetValue("outFile", out s))
                if (!s.Equals(""))
                    message += "\"" + s + "\" ";
            else
                message += "stdout "; // use stdout
            if (opts.TryGetValue("dataFile", out s))
                if (!s.Equals(""))
                    message += " -data \"" + s + "\" ";
            if (opts.TryGetValue("keyCode", out s))
                message += " -kc " + s;
            if (opts.TryGetValue("keyName", out s))
                message += " -kn " + s;
            if (opts.TryGetValue("guiOff", out s))
                message += " -guioff";
            if (opts.TryGetValue("comp15", out s))
                message += " -comp15";
            if (opts.TryGetValue("owner", out s))
                message += " -o \"" + s + "\"";
            if (opts.TryGetValue("user", out s))
                message += " -u \"" + s + "\"";
            if (opts.TryGetValue("noAnnote", out s))
                message += " -noannote";
            if (opts.TryGetValue("noCopy", out s))
                message += " -nocopy";
            if (opts.TryGetValue("noPrint", out s))
                message += " -noprint";
            if (opts.TryGetValue("noChange", out s))
                message += " -nochange";
            if (opts.TryGetValue("enc128", out s))
                message += " -e128";
            if (opts.TryGetValue("aes", out s))
                message += " -aes " + s;
            if (opts.TryGetValue("extOpts", out s))
                message += " " + s;
            if (opts.TryGetValue("allowBreaks", out s))
                message += " -allowbreaks";
            if (opts.TryGetValue("quick", out s))
                message += " -q";
            if (opts.TryGetValue("quick2", out s))
                message += " -q2";
            if (opts.TryGetValue("open", out s))
                message += " -open";            
            if (opts.TryGetValue("print", out s))
                message += " -print";            
            if (opts.TryGetValue("autoSend", out s))
                message += " -autosend";            
            if (opts.TryGetValue("licInfo_name", out s)){
                String p = "";
                String d = "";
                opts.TryGetValue("licInfo_name", out p);
                opts.TryGetValue("licInfo_autodl", out d);
                message += " -licname " + s + " -licpwd " + p + (d.Equals("1") ? " -licweb" : "");
            }
            return message;
        }
    
        // Send the BUILDPDF command to server to run the commands
        private (byte[], String) callTCP(bool isStop = false,
            bool isStatus = false,
            bool retBytes = false,
            String saveFile = ""){
            String errMsg = "";
            Byte[] data = {};
            Byte[] bytes = {};
            Object host = "";
            bool retPDF = retBytes;
            // String to store the response ASCII representation.
            String responseData = String.Empty;

            if (!server.TryGetValue("host", out host)){
                return (bytes, "Server not setup - try calling setServer or startServer first.");
            }

            if (!client.Connected) {
                client = new TcpClient((String) host, (int) server["port"]);
                stream = client.GetStream();
            }

            errMsg = sendTCP();
            if (errMsg.Equals("")){
                if (opts.TryGetValue("autoSend", out String s))
                  retPDF = true; // need to keep open and send files
            }
            if (!saveFile.Equals("")){
                retPDF = true; // need to save the PDF
            }

            try {         
                if (retPDF){
                    data = System.Text.Encoding.ASCII.GetBytes(" -return ");             
                    stream.Write(data, 0, data.Length);      
                }
                data = System.Text.Encoding.ASCII.GetBytes("\nBUILDPDF\n");             
                stream.Write(data, 0, data.Length);      
                if (isStatus){ 
                    do {
                        data = new Byte[1024];
                        // Read the first batch of the TcpServer response bytes.
                        Int32 rawData = stream.Read(data, 0, data.Length);
                        responseData += System.Text.Encoding.ASCII.GetString(data, 0, rawData);
                    }
                    while (stream.DataAvailable);
                } else if (!isStop) {
                    MemoryStream memstream = new MemoryStream();
                    Socket s = client.Client;
                    byte[] buffer = new byte[1024];                    
                    int bytesRead = 0;
                    String retStr = "";
                    String[] retCmds = new String[2];                    
                    
                    while(true) {
                        // Read the first batch of the TcpServer response bytes.
                        bytesRead = stream.Read(buffer, 0, buffer.Length);
                        retStr = System.Text.Encoding.ASCII.GetString(buffer, 0, buffer.Length);
                        if (retStr.ToLower().StartsWith("content-length:")){
                          // Receiving the PDF back
                          while ((bytesRead = stream.Read(buffer, 0, buffer.Length)) > 0){
                              memstream.Write(buffer, 0, bytesRead);     
                          }                        
                          bytes = memstream.ToArray();
                          memstream.SetLength(0);
                          break;
                        }
                        else
                        {
                          retCmds = retStr.Split(':');                           
                          if (retCmds.Length > 0 && 
                              (retCmds[0].ToLower().Equals("send-file")
                              || retCmds[0].ToLower().Equals("send-md5"))) {

                              if (retCmds[0].ToLower().Equals("send-file")){
                              FileInfo f = new FileInfo(retCmds[1]);
                              data = System.Text.Encoding.ASCII.GetBytes("Content-Length: " + f.Length + "\n\n");
                              stream.Write(data, 0, data.Length);
                              BinaryReader br = new BinaryReader(new FileStream(retCmds[1], FileMode.Open));

                              while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                                  stream.Write(buffer, 0, bytesRead);     
                              stream.Flush();                                       
                              } else {
                                  String message = "";
                                  MD5 md5 = MD5.Create();
                                  FileStream stream = File.OpenRead(retCmds[1]);
                                  byte[] md5Bytes = md5.ComputeHash(stream);
                                  String fileHash = ByteArrayToString(md5Bytes);
                                  message = "Content-Length: " + fileHash.Length + "\n\n" + fileHash;
                                  data = System.Text.Encoding.ASCII.GetBytes(message);             
                                  stream.Write(data, 0, data.Length);
                                  stream.Flush();
                              }
                          }
                          if ((s.Poll(1000, SelectMode.SelectRead) && (s.Available == 0)) || !s.Connected)
                          {
                              break;
                          } 
                        }                     
                    }
                }                
                stream.Close();
                client.Close();
                if (!saveFile.Equals("")){
                    File.WriteAllBytes (saveFile, bytes);
                    Array.Clear(bytes, 0, bytes.Length);
                }
                if (!retBytes){
                    Array.Copy(bytes, new byte[0], 0);
                }
            } catch (SocketException e) {
                errMsg = e.Message;
            }
            return (bytes, (responseData.Equals("") ? errMsg : responseData));
        }

        private static string ByteArrayToString(byte[] ba)
        {
            return BitConverter.ToString(ba).Replace("-","");
        }
    }
}
