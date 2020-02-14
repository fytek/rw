<?php
$pdfobj = new COM('FyTek.ReportWriter');

# Normally this is done in a different program or shell script.
# Also, you might want to start a server on a different box altogether.
$pdfobj->startServer();

# This is just one way to send commands.
# You may create a file on disk as well and the file instead.
$pdfobj->setPDFCmd ("<PDF>");
$pdfobj->setPDFCmd ("<PAGE>");
$pdfobj->setPDFCmd ("<TEXT ALIGN=C>");
$pdfobj->setPDFCmd ("Hello, world");
$pdfobj->setPDFCmd ("</TEXT>");
$pdf = $pdfobj->buildReport(true,'c:\temp\hello.pdf'); # this is the output file
print $pdf->Msg; # if all went well should return OK

# Normally you would leave the server running for the next
# request but since this is just a sample we'll stop it.
$pdfobj->stopServer();
?>
