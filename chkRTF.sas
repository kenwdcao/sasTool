/**********************************************************************************************
    Program Name: chkRTF.sas
         @Author: Ken Cao (yong.cao@ppdi.com)
         @Initial Date: 2016/07/13
   
 
    Parameter:
         RTFNAME: Optional. Postional parameter. File name of the RTF file to be checked. 
                  If omitted, this program will check all RTFs under directory RTFDIR.
          RTFDIR: Required. Keyword parameter. File path of RTF file(s) to be checked. 
    

    Description:
        Macro chkRTF is designed to check pagination issue of RTF Table/Listing generated
        by SAS. The macro behaves differently depending on how this macro was called:

        1) Parameter RTFNAME is specified. Macro will check against specified RTF and prints
           paginiations in SAS log. 

            %chkRTF(RTF-NAME.rtf, rtfdir=TLF-DIRECTORY)
              
           If pagination issue was detected, macro will print a ALERT_P message alerting user 
           of page break issue with another message that displays the first occurrence of page 
           break issue.        
           
        2) Parameter RTFNAME is't omitted. Macro will check against all RTF files in the
           directory of RTFDIR and prints paginations in an Excel file. The excel file is
           saved as _pageChk.xls in the same folder where this macro was called.
           
           %chkRTF(rtfdir=TLF-DIRECTORY)
            
***********************************************************************************************/

%macro chkRTF(rtfname, rtfdir=);

%local line_size;    %let line_size = %sysfunc(getoption(ls));
%local dash_line;    %let dash_line  = %sysfunc(repeat(_, %eval(&line_size-10)));
%local equal_line;   %let equal_line = %sysfunc(repeat(=, %eval(&line_size-10)));

option noxwait xsync ls=140;

%put ALERT_I: ;
%put ALERT_I: &equal_line;
%put ALERT_I: Macro &sysmacroname STARTs...;

%local thisdir;
%local workdir;
%local mode; /* 1=check single RTF 2=check folder */
%local vbsrc;
%local ls;

** Mask potential special characters in paramters;
%let rtfdir = %superq(rtfdir);
%let rtfname = %superq(rtfname);

** get current direcotry path and SAS work library path;
libname _here '.';
%let thisdir = %qsysfunc(pathname(_here));
%let workdir = %qsysfunc(pathname(work));
libname _here clear;


************************************************************************;
* Parameter check
************************************************************************;

%if %qsubstr(&rtfdir, %length(&rtfdir)) = %str(/) or %qsubstr(&rtfdir, %length(&rtfdir)) = %str(\) %then %do;
    %let rtfdir = %qsubstr(&rtfdir, 1, %length(&rtfdir)-1); ** trim the ending slash;
%end;

%if %length(&rtfname) > 0 %then %do;
    %if %sysfunc(fileexist(&rtfdir\&rtfname)) = 0 %then %do;
        %put ALERT_P(&sysmacroname): RTFNAME is specified as &rtfname but file was not found.;
        %goto MACROEND;
    %end;
%end;

%if %length(&rtfname) > 0 %then %let mode = 1;
%else %let mode = 2;


************************************************************************;
* Generate VBScript (a VBScript function and Sub)
* The DATA step is very long since I created varaible line/len to maintan
* leading space in the script in case I want to debug it
************************************************************************;
%put NOTE-Generating VBScript &workdir\_chkRTF.vbs;
data _null_;
  length line $255 len 8;
  file "&workdir\_chkRTF.vbs" lrecl=255;
  put " ";
  line = "'Function to return array of # of section/# of pages/# of first occurrence of pagination issue";
  len = length(line);
  put line $varying1024. len;
  line = "Function chkRTF_KNL(objWORD, sRTFDIR, sRTFNAME)";
  len = length(line);
  put line $varying1024. len;
  line = "  Dim rtf";
  len = length(line);
  put line $varying1024. len;
  line = "  Dim rtnArr(4)";
  len = length(line);
  put line $varying1024. len;
  line = "  Const wdActiveEndPageNumber = 3 'Active End Page Number (of a section)";
  len = length(line);
  put line $varying1024. len;
  line = "  ";
  len = length(line);
  put line $varying1024. len;
  line = "  nsection = 0";
  len = length(line);
  put line $varying1024. len;
  line = "  npage = 0";
  len = length(line);
  put line $varying1024. len;
  line = "  fpagen = 0";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "  Set rtf = objWORD.Documents.Open(sRTFDIR & ""\"" & sRTFNAME, False ,True)";
  len = length(line);
  put line $varying1024. len;
  line = "  ";
  len = length(line);
  put line $varying1024. len;
  line = "  nsection = rtf.Sections.Count";
  len = length(line);
  put line $varying1024. len;
  line = "  rtf.Repaginate";
  len = length(line);
  put line $varying1024. len;
  line = "  npage = rtf.Sections(rtf.Sections.Count).Range.Information(wdActiveEndPageNumber)";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "  if npage <> nsection Then";
  len = length(line);
  put line $varying1024. len;
  line = "      For Each sec in rtf.Sections";
  len = length(line);
  put line $varying1024. len;
  line = "          nEndPgn = sec.Range.Information(wdActiveEndPageNumber)";
  len = length(line);
  put line $varying1024. len;
  line = "          If nEndPgn <> sec.Index Then";
  len = length(line);
  put line $varying1024. len;
  line = "              fpagen = nEndPgn";
  len = length(line);
  put line $varying1024. len;
  line = "                fsect = sec.Index";
  len = length(line);
  put line $varying1024. len;
  line = "              Exit For";
  len = length(line);
  put line $varying1024. len;
  line = "          End If";
  len = length(line);
  put line $varying1024. len;
  line = "      Next";
  len = length(line);
  put line $varying1024. len;
  line = "  End If";
  len = length(line);
  put line $varying1024. len;
  line = "  ";
  len = length(line);
  put line $varying1024. len;
  line = "  rtnArr(0) = nsection";
  len = length(line);
  put line $varying1024. len;
  line = "  rtnArr(1) = npage";
  len = length(line);
  put line $varying1024. len;
  line = "  rtnArr(2) = fpagen";
  len = length(line);
  put line $varying1024. len;
  line = "    rtnArr(3) = fsect";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "  rtf.Close(False)";
  len = length(line);
  put line $varying1024. len;
  line = "  ";
  len = length(line);
  put line $varying1024. len;
  line = "  chkRTF_KNL = rtnArr";
  len = length(line);
  put line $varying1024. len;
  line = "End Function";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "Sub chkRTF(sRTFDIR, sRTFNAME, sRPTDIR)";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "    Dim objWrd";
  len = length(line);
  put line $varying1024. len;
  line = "    Dim objExl";
  len = length(line);
  put line $varying1024. len;
  line = "    Dim objFSO";
  len = length(line);
  put line $varying1024. len;
  line = "    Dim objRpt";
  len = length(line);
  put line $varying1024. len;
  line = "    Dim fld";
  len = length(line);
  put line $varying1024. len;
  line = "    Const sRptFileName = ""_pageChk""";
  len = length(line);
  put line $varying1024. len;
  line = "    Const xlMaximized = -4137";
  len = length(line);
  put line $varying1024. len;
  line = "    Const xlLeft = -4131";
  len = length(line);
  put line $varying1024. len;
  line = "    Const XlRight = -4152";
  len = length(line);
  put line $varying1024. len;
  line = "    Const xlWorkbookNormal = -4143";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "    Set objFSO=CreateObject(""Scripting.FileSystemObject"")";
  len = length(line);
  put line $varying1024. len;
  line = "    Set objWrd = CreateObject(""Word.Application"")";
  len = length(line);
  put line $varying1024. len;
  line = "    objWrd.Visible = False";
  len = length(line);
  put line $varying1024. len;
  line = "    OpenAttachmentsInFullScreen = objWrd.OpenAttachmentsInFullScreen";
  len = length(line);
  put line $varying1024. len;
  line = "    objWrd.OpenAttachmentsInFullScreen = False";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "    If Not IsNull(sRTFNAME) Then";
  len = length(line);
  put line $varying1024. len;
  line = "        ' If checking single file then use a text file to record pagination";
  len = length(line);
  put line $varying1024. len;
  line = "        pagearr = chkRTF_KNL(objWrd, sRTFDIR, sRTFNAME)";
  len = length(line);
  put line $varying1024. len;
  line = "        Set objRpt = objFSO.CreateTextFile(sRPTDIR & ""\"" & sRptFileName & "".txt"", True)";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Write pagearr(0) & ""|"" & pagearr(1) & ""|"" & pagearr(2)";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Close";
  len = length(line);
  put line $varying1024. len;
  line = "    Else";
  len = length(line);
  put line $varying1024. len;
  line = "        ' Start an Excel to record paginations";
  len = length(line);
  put line $varying1024. len;
  line = "        Set objExl = CreateObject(""Excel.Application"")";
  len = length(line);
  put line $varying1024. len;
  line = "        objExl.Visible = True";
  len = length(line);
  put line $varying1024. len;
  line = "        objExl.DisplayAlerts = False";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "        Set fld = objFSO.GetFolder(sRTFDIR)";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "        ' # of RTF files";
  len = length(line);
  put line $varying1024. len;
  line = "        nrtf = 0";
  len = length(line);
  put line $varying1024. len;
  line = "        For Each file In fld.Files";
  len = length(line);
  put line $varying1024. len;
  line = "            If InStr(UCase(file.name), "".RTF"") and Left(file.name,2) <> ""~$"" Then";
  len = length(line);
  put line $varying1024. len;
  line = "                nrtf = nrtf + 1";
  len = length(line);
  put line $varying1024. len;
  line = "            End If";
  len = length(line);
  put line $varying1024. len;
  line = "        Next";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "        Set objRpt = objExl.workbooks.add";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Windows(1).WindowState = xlMaximized";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "        repshtname = ""Pagenition""";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Activesheet.name = repshtname";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "        ' Summary Information";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Cells(1, 1) = ""RTF directory:""";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Cells(1, 2) = fld.path";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Range(""B1:Z1"").Merge";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Range(""B1:Z1"").HorizontalAlignment = xlLeft";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Cells(2, 1) = ""# of Processed:""";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Cells(2, 2) = ""=concatenate(counta(A6:A65536), """" / """", "" & nrtf &  "")""";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Range(""A2:B2"").font.color = RGB(0, 176, 80)";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Cells(2, 2).HorizontalAlignment = XlRight";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Cells(3, 1) = ""# of Suspicious:""";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Cells(3, 2) = ""=counta(D6:D65536)""";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Range(""A3:B3"").font.color = RGB(255, 0, 112)";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "        ' Headers (starting at line 5)";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Cells(5, 1) = ""File Name""";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Cells(5, 2) = ""Logical Page #""";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Cells(5, 3) = ""Physical Page #""";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Cells(5, 4) = ""Check Page (around)""";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Range(""A5:D5"").Interior.color = RGB(112, 48, 160)";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Range(""A5:D5"").font.color = RGB(255, 255, 255)";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Range(""A5:D5"").font.bold = True";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Range(""A5:D5"").AutoFilter";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "        objRpt.Worksheets(repshtname).Columns(1).AutoFit";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Columns(2).AutoFit";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Columns(3).AutoFit";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Columns(4).AutoFit";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "        objRpt.Worksheets(repshtname).Rows(6).Select";
  len = length(line);
  put line $varying1024. len;
  line = "        objExl.ActiveWindow.FreezePanes = True";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "        rtfIndex = 0";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "        ' Walking the directory";
  len = length(line);
  put line $varying1024. len;
  line = "        For Each file In fld.Files";
  len = length(line);
  put line $varying1024. len;
  line = "            If InStr(UCase(file.name), "".RTF"") and Left(file.name,2) <> ""~$"" Then";
  len = length(line);
  put line $varying1024. len;
  line = "                rtfIndex = rtfIndex + 1";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "                objRpt.Worksheets(repshtname).Rows(rtfIndex+5).Select 'move to current cell being written in";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "                objRpt.Worksheets(repshtname).Cells(rtfIndex+5, 1) = file.name";
  len = length(line);
  put line $varying1024. len;
  line = "                objRpt.Worksheets(repshtname).Columns(1).AutoFit";
  len = length(line);
  put line $varying1024. len;
  line = "                objRpt.Worksheets(repshtname).Cells(rtfIndex+5, 2) = ""Opening..."" ' status reporting";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "                pagearr = chkRTF_KNL(objWrd, sRTFDIR, file.name)";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "                objRpt.Worksheets(repshtname).Cells(rtfIndex+5, 2) = pagearr(0)";
  len = length(line);
  put line $varying1024. len;
  line = "                objRpt.Worksheets(repshtname).Cells(rtfIndex+5, 3) = pagearr(1)";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "                If pagearr(2) > 0 Then";
  len = length(line);
  put line $varying1024. len;
  line = "                    If pagearr(3) = 1 Then";
  len = length(line);
  put line $varying1024. len;
  line = "                        objRpt.Worksheets(repshtname).Cells(rtfIndex+5, 1) = ""=HYPERLINK("""""" & fld.path & ""\"" & file.name & """""", """""" & file.name & """""")""";
  len = length(line);
  put line $varying1024. len;
  line = "                    Else";
  len = length(line);
  put line $varying1024. len;
  line = "                        objRpt.Worksheets(repshtname).Cells(rtfIndex+5, 1) = ""=HYPERLINK("""""" & fld.path & ""\"" & file.name & ""#IDX"" & pagearr(3) - 1 & """""", """""" & file.name & """""")""";
  len = length(line);
  put line $varying1024. len;
  line = "                    End If";
  len = length(line);
  put line $varying1024. len;
  line = "                    objRpt.Worksheets(repshtname).Range(""A"" & rtfIndex+5 & "":D"" & rtfIndex+5).Interior.color = RGB(255, 153, 153)";
  len = length(line);
  put line $varying1024. len;
  line = "                    objRpt.Worksheets(repshtname).Cells(rtfIndex+5, 4) = pagearr(2)";
  len = length(line);
  put line $varying1024. len;
  line = "                End If";
  len = length(line);
  put line $varying1024. len;
  line = "            End If";
  len = length(line);
  put line $varying1024. len;
  line = "        Next";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "        ' -4143 = Normal Workbook";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.Worksheets(repshtname).Rows(6).Select";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.SaveAs sRPTDIR & ""\"" & sRptFileName & "".xls"", xlWorkbookNormal";
  len = length(line);
  put line $varying1024. len;
  line = "        objRpt.close";
  len = length(line);
  put line $varying1024. len;
  line = "        objExl.Quit";
  len = length(line);
  put line $varying1024. len;
  line = "    End If";
  len = length(line);
  put line $varying1024. len;
  put " ";
  line = "    objWrd.OpenAttachmentsInFullScreen = OpenAttachmentsInFullScreen";
  len = length(line);
  put line $varying1024. len;
  line = "    objWrd.Quit(False)";
  len = length(line);
  put line $varying1024. len;
  line = "End Sub";
  len = length(line);
  put line $varying1024. len;
run;

** Now call VBScript Sub chkRTF according to MODE;
data _null_;
  length line $1024 _rtf $255;
  file "&workdir\_chkRTF.vbs" lrecl=1024 mod;
  %if &mode = 1 %then %do;
    _rtf = """&rtfname""";  
    _rptdir = """&workdir""";
  %end;
  %else %if &mode = 2 %then %do;
    _rtf = "null";
    _rptdir = """&thisdir""";
  %end;
  line = "chkRTF ""&rtfdir"", "||trim(_rtf)||', '||_rptdir;
  put line;
  putlog 'NOTE-Executing VBScript Sub as: ' line=;
run;


************************************************************************;
* Executing Script 
* If somehow execution failed, printing script in SAS log
************************************************************************;
data _null_;
    rc = system("""&workdir\_chkRTF.vbs""");
    put "NOTE-Exit code after executing VBScript: " rc;
    CALL SYMPUT('vbsrc', strip(put(rc, best.)));
run;

%if &vbsrc > 0 %then %do;
    %put ALERT_P(&sysmacroname): VBScript generated but execution was failed. Return code: &vbsrc;
    %goto MACROEND;
%end;
%else %if &mode = 2 %then %goto MACROEND;


************************************************************************;
* Analyze result file (mode 1)
************************************************************************;
data _null_;
    length logicPageNum phyPageNum _1stPageIssue 8;
    infile "&workdir\_pageChk.txt" lrecl=255 truncover;
    input;
    logicPageNum = input(scan(_infile_, 1, '|'), best.);
    phyPageNum = input(scan(_infile_, 2, '|'), best.);
    _1stPageIssue = input(scan(_infile_, 3, '|'), best.);
    
    put "ALERT_I: Checking RTF file: &rtfdir\&rtfname..";
    put "ALERT_I: Logical Page Number : " logicPageNum;
    put "ALERT_I: Physical Page Number: " phyPageNum;
    if logicPageNum ^= phyPageNum then do;
        put "ALERT_P(&sysmacroname): pagination issue detected. Please check page " _1stPageIssue '.';
    end;
    else put "ALERT_I: No pagination issue found"; 
run;


************************************************************************;
* Program End 
************************************************************************;
%MACROEND:
    option ls=&line_size; ** reset LS option to user default value;
    %put ALERT_I: Macro &sysmacroname ENDs...;
    %put ALERT_I: &equal_line;

%mend chkRTF;
