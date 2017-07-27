/********************************************************************************
* Program Name: strSearch.sas
*   @Author: Ken Cao (Yong.Cao@ppdi.com)
*   @Initial Date: 2015/10/09
*
* ##############################################################################
* This program is for searching strings from text files under a folder. It supports
* to search up to 10 strings simutaneously. User can specify whethr or not the 
* search string is case senstive (no by default) and regular expression search 
* independently for each of the search string. The output is a traditional .lst
* ouptut which contains the file name that has the postive match with rows matched.
* 
* The key inputs of this macro are:
*   Folder path where files to be searched reside in 
*   Whether or not to search recursively for all files under the subfolders
*   The target file extension
*   The target file name pattern (in regular expression)
*   The search string(s)
*
* Since this program uses windows command line and pipe to retrieve the target 
* files, it will only work for SAS in Windows. But porting this macro to UNIX
* platform shouldn't be difficult.
*
* ##############################################################################
* PARAMETERS:
*   DIR: Folder path where files to be searched reside in.
*   ----------------------------------------------------------------------------
*   INCLUDESUBFOLDER: Whether or not to search recursively under DIR.
*   ----------------------------------------------------------------------------
*   FILENAMEPATTERN: The target file name pattern in regular expression. For 
*     example if you need to search all files that start with V, specify ^v. Note
*     this is case insenstive.
*   ----------------------------------------------------------------------------
*   FILETYPE: The extension name of the target file. Case insensitive.
*   ----------------------------------------------------------------------------
*   RECURSIVE: Whether or not to perform recursive search. A recursive search is,
*     when file A matches any search string, and file B contains the name of file
*     A then file B is included in the report.
*   ----------------------------------------------------------------------------
*   SEARCHSTR1-10: Target string to be searched among target files.
*   ----------------------------------------------------------------------------
*   EXCLUDESTR1-10: Strings to be excluded. This is used in pair with SEARCHSTR.
*     For example, if you specify '"ATC\dCD" Y' for SEARCHSTR1 then it will match 
*     strings like "ATC1CD", "ATCD2CD" ... Suppose that you don't want to match 
*     with ATC3CD then you can specify EXCLUESTR1="ATC3CD". This is a bad example
*     as there is smarter way to do this but you should get the idea. See note 
*     below for the syntax of search string and exclude string.
*   
* ##############################################################################
* Syntax of search string (and exclusion string)
* 1. Specify a normal text string. Suppose you want to search "TRT01AN", then 
*  all you need is to pass that string along with the quotation mark to paramter
*  SEARCHSTR1. 
*   ----------------------------------------------------------------------------
* 2. Specify a regular expression search string. Suppose you want to search all
*  possible treatment variables, you can pass on '"TRT\d+(A|N)" Y'. Note the outer
*  single quotation mark should not be included. Also note that the ending "Y" 
*  specifies this is a regular expression search. This is a switch that defaults
*  to N.
*   ----------------------------------------------------------------------------
* 3. Specify case sensitivity. Suppose you want to match the string "Time" in the 
*  exact case. Specify it as '"Time" N Y'. The outer single quotation mark is not
*  the part of it. The first N specifies that this is not a regular expression 
*  and the last Y specifies that the serach is case senstive. This swtich is 
*  defaulted to N (case insensitive).
*
* ##############################################################################
* Example call
* The following macro invcation searches SAS programs of which file name starts
* with "V" and the target string is compress function call.
*
* %strSearch( 
*   dir              = C:\test\
*  ,filenamePattern  = ^v
*  ,filetype         = sas
*  ,includeSubFolder = N
*  ,searchStr1       = "compress\(" Y 
* );
*
********************************************************************************/


%macro parse_searchstr(searchStr_in, mvar_searchStr_out, mvar_reg_exp, mvar_case_sens);

%local pattern;

/* pattern of search string in regular expression */
%let pattern = ^([''""])(.+\1)(\s+[Y|N])?(\s+[Y|N])?$; 

%local ___searchStr;
%local ___regExp;
%local ___caseSense;

%let searchStr_in = %sysfunc(strip(&searchStr_in));

%if %sysfunc(prxmatch(/&pattern/i, &searchStr_in)) = 0 %then %do;
  %put ERR%str(OR): Invalid search string;
  %return;
%end;


%let ___searchStr = %sysfunc(prxchange(s/&pattern/$1$2/i, -1, &searchStr_in));
%let ___regExp = %upcase(%sysfunc(prxchange(s/&pattern/$3/i, -1, &searchStr_in)));
%let ___caseSense = %upcase(%sysfunc(prxchange(s/&pattern/$4/i, -1, &searchStr_in)));

%if %length(&___regExp) = 0 %then %let ___regExp = N;
%if %length(&___caseSense) = 0 %then %let ___caseSense = N;


%if &___regExp = Y %then %do;
  %if &___caseSense = N %then %do;
    %let ___searchStr = %sysfunc(prxchange(s/([''""])(.*)\1/$1\/$2\/i$1/, -1, &___searchStr));
  %end;
  %else %do;
    %let ___searchStr = %sysfunc(prxchange(s/([''""])(.*)\1/$1\/$2\/$1/, -1, &___searchStr));
  %end; 
%end;

%let &mvar_searchStr_out = &___searchStr;
%let &mvar_reg_exp = &___regExp;
%let &mvar_case_sens = &___caseSense;

%mend parse_searchstr;



%macro strSearchKNL(
  infile             =
 ,searchStr          =
 ,search_case_sense  =
 ,search_reg_exp     =
 ,excludeStr         =
 ,exclude_case_sense =
 ,exclude_reg_exp    =
);

  %local rc;
  %local filrf;
  %local seachStrPRX;

  %let filrf=%str(__search);
  %let rc=%sysfunc(filename(filrf,&infile));


  data _matchFile;
    keep directory file line filewpath linenum ;
    len=1024;
    infile &filrf lrecl=32767 truncover;
    input _line_ $char32767. ;
    char=substr(_line_,1,1);
    linenum+1;

    length line2  $32767;
    *replace tab delimiter with normal blank;
    line2=prxchange('s/\s/ /',-1,strip(_line_));
    line2=strip(line2);

    length line $256 directory filewpath $1024 file $256;
    file=scan("&infile",-1,"\/");
    filewpath="&infile";
    directory=substr("&infile",1,length(strip("&infile"))-length(file)-1);

    %if &search_reg_exp = Y %then %do;
      if prxmatch(&searchStr, strip(_line_)) = 0 then return;
    %end;
    %else %do;
      %if &search_case_sense = Y %then %do;
        if index(strip(_line_), &searchStr) = 0 then return;
      %end;
      %else %do;
        if index(upcase(strip(_line_)), %upcase(&searchStr)) = 0 then return;
      %end;
    %end;
    
    %if %length(&excludeStr) > 0 %then %do;
      %if &exclude_reg_exp = Y %then %do;
        if prxmatch(&excludeStr, strip(_line_)) then return;
      %end;
      %else %do;
        %if &exclude_case_sense = Y %then %do;
          if index(strip(_line_), "&excludeStr") then return;
        %end;
        %else %do;
          if index(upcase(strip(_line_)), %upcase(&excludeStr)) then return;
        %end;
      %end;
    %end;

    line = substr(line2, 1, 256);
    output;
  
  run;

  %let rc=%sysfunc(filename(filrf));
  
%mend strSearchKNL;

%macro strSearch( dir=
                 ,filenamePattern = 
                 ,filetype =
                 ,includeSubFolder = N
                 ,recursive = N
                 ,searchStr  =, excludeStr  =
                 ,searchStr1 =, excludeStr1 =
                 ,searchStr2 =, excludeStr2 =
                 ,searchStr3 =, excludeStr3 =
                 ,searchStr4 =, excludeStr4 =
                 ,searchStr5 =, excludeStr5 =
                 ,searchStr6 =, excludeStr6 =
                 ,searchStr7 =, excludeStr7 =
                 ,searchStr8 =, excludeStr8 =
                 ,searchStr9 =, excludeStr9 =
                 ,searchStr10=, excludeStr10=
);
   
%local workdir;
%local nfiles;
%local rc;
%local dsid;
%local nobs;
%local i;
%local j;
%local file;
%local nfound;
%local rundt;

%let nfound=0;

%let workdir = %qsysfunc(pathname(work));

%let recursive = %upcase(&recursive);
%let includeSubFolder = %upcase(&includeSubFolder);

%if %length(&recursive) = 0 %then %let recursive = N;
%else %let recursive = %substr(&recursive, 1, 1);

%if %length(&includeSubFolder) = 0 %then %let includeSubFolder = N;
%else %let includeSubFolder = %substr(&includeSubFolder, 1, 1);

*************************************************************************************;
* Get a list of file names according to user input                                   ;
*************************************************************************************;
option noxwait xsync;

data _null_;
  length cmd1 cmd2 $1024;
  %if &includeSubFolder = N %then %do;
  cmd1 = "dir /b ""&dir"" > ""&workdir\list.txt""";
  cmd2 = "dir /b /ad ""&dir"" > ""&workdir\excl.txt""";
  %end;
  %else %do;
  cmd1 = "dir /b /s ""&dir"" > ""&workdir\list.txt""";
  cmd2 = "dir /b /ad /s ""&dir"" > ""&workdir\excl.txt""";
  %end;
  rc = system(cmd1); ** get all child-items;
  rc = system(cmd2); ** get sub-directories;
run;

data child;
  length item $1024;
  infile "&workdir\list.txt" lrecl=1024 truncover;
  input item $1024.;
run;

data subdir;
  length item $1024;
  infile "&workdir\excl.txt" lrecl=1024 truncover;
  input item $1024.;
run;

proc sort data=child; by item; run;
proc sort data=subdir; by item; run;

data files;
  length filename $255 directory $1024 filetype $255;
  merge child
        subdir(in=__dir)
  ;by item;
  if not __dir;
  %if &includeSubFolder = N %then %do;
  item = strip("&dir")||'\'||item;
  %end; 
  filename = scan(item, -1, '/\');
  if index(filename, '.') then filetype = scan(filename, -1, '.');
  directory = substr(item, 1, length(item)-length(filename)-1);

  keep directory filename filetype;
run;

*************************************************************************************;
* In case of filetype is specified, macro getFileNames returns file names of which   ;
* file type suffix containing user input (filetype). But in this macro, we only need ;
* file of which file type suffix precisely matched user input.                       ;
* If user specified filenamePatttern, it cannot be processed by macro getFileNames,  ;
* this filter is done below.                                                         ;
*************************************************************************************;
proc sort data=files;
  by directory filename;
  where 1 /* place holder for where statement */
  %if %length(&filetype) > 0 and "&filetype" ^= "*" %then %do;
  and upcase(filetype)="%upcase(&filetype)"
  %end;
  %if %length(&filenamePattern) > 0 %then %do;
  %let filenamePattern = %sysfunc(prxchange(s/[""]/""/, -1, &filenamePattern));
  /* on windows, file name is case-insensitive */
  and prxmatch("/&filenamePattern/i", strip(filename)) 
  %end;
  ;
run;


%let dsid=%sysfunc(open(files));
%let nfiles=%sysfunc(attrn(&dsid,nobs));
%let rc=%sysfunc(close(&dsid));

** master dataset for search result;
data matchFiles;
  length line file $256 filewpath directory $1024 linenum 8;
  label filewpath = 'File';
  call missing(line,file,filewpath,directory,linenum);
  if 0;
run;


%if %length(&searchStr) > 0 and %length(&searchStr1) = 0 %then %do;
  %let searchStr1 = &searchStr;
%end;
%if %length(&excludeStr) > 0 and %length(&excludeStr1) = 0 %then %do;
  %let excludeStr1 = &excludeStr;
%end;

%do i = 1 %to 10;
  %local _searchStr&i;
  %local _s_case&i;
  %local _s_reg&i;
  %local _excludeStr&i;
  %local _e_case&i;
  %local _e_reg&i;

  %if %length(&&searchStr&i) > 0 %then %do;
    %parse_searchstr(&&searchStr&i, _searchStr&i, _s_reg&i, _s_case&i);
  %end;
  %if %length(&&excludeStr&i) > 0 %then %do;
    %parse_searchstr(&&excludeStr&i, _excludeStr&i, _e_reg&i, _e_case&i);
  %end;
%end;


%do i=1 %to &nFiles;
  data _null_;
    set files (firstobs=&i obs=&i);
    call symput('file', strip(directory)||'\'||strip(filename));
  run;
  %do j = 1 %to 10;
    %if %length(&&searchStr&j) > 0 %then %do;
    ***************************************************************************;
    * Macro strSearchKNL returns search results in a dataset called _matchFile ;
    * who has same structure as master dataset matchFiles.                     ;
    ***************************************************************************;
      %strSearchKNL(
        infile             = &file
       ,searchStr          = &&_searchStr&j
       ,search_case_sense  = &&_s_case&j
       ,search_reg_exp     = &&_s_reg&j
       ,excludeStr         = &&_excludeStr&j
       ,exclude_case_sense = &&_e_case&j
       ,exclude_reg_exp    = &&_e_reg&j
       );
      proc append base=matchFiles data=_matchFile; run;
    %end;
  %end;
%end;

***************************************************************************;
* In case of recursive search -  the matched file name is searched         ;
***************************************************************************;
%if &recursive = Y %then %do;
  %local _nFilesNew;  
  %local _sFile;
  %let _newFiles = 0;

  proc sort data=matchFiles nodupkey out=_newFiles(keep=file); by file; run;
  proc sort data=matchFiles nodupkey out=_oldFiles(keep=file); by file; run;
  proc sql noprint;
    select count(*)
    into: _nFilesNew
    from _newFiles
    ;
  quit;

  %do %while(&_nFilesNew > 0);
    %do i = 1 %to &nFiles;
      data _null_;
        set files (firstobs=&i obs=&i);
        call symput('file', strip(directory)||'\'||strip(filename));
      run;
      %do j = 1 %to &_nFilesNew.;
        data _null_;
          set _newFiles(firstobs=&j obs=&j);
          call symput('_sFile', strip(file));
        run;
        %strSearchKNL(
          infile             = &file
         ,searchStr          = "&_sFile"
         ,search_case_sense  = N
         ,search_reg_exp     = N
         );

        ** No need to include self;
        proc sql;
          delete *
          from _matchFile
          where upcase(file) = %upcase("&_sFile")
          ;
        quit;
        proc append base=matchFiles data=_matchFile; run;
      %end;
    %end;
    proc sort data=matchFiles nodupkey out=_mFilesNew(keep=file); by file; run;
    data _newFiles;
      merge _oldFiles  (in=__old)
            _mFilesNew (in=__new) 
      ;by file;
      if not __old;
    run;
    proc sort data=matchFiles nodupkey out=_oldFiles(keep=file); by file; run;

    %let _nFilesNew = 0;
    proc sql noprint;
      select count(*)
      into: _nFilesNew
      from _newFiles
      ;
    quit;
  %end;
%end;

proc sort data=matchFiles nodupkey; 
  by filewpath linenum;
run;

proc sql noprint;
  select count(distinct filewpath)
  into: nFound
  from matchFiles
  ;
quit;

%let nFound = %sysfunc(strip(&nFound));



/*Run data/time*/
data _null_;
  datetime=put(datetime(),is8601dt19.);
  call symputx('rundt',datetime);
run;

***************************************************************************;
* Create a summary header in the search result output file.                ;
***************************************************************************;

 data _title;      
   length title $256;
   title="Root Directory: &dir"; output;
   title="Include Subfolder: %upcase(&includeSubFolder)"; output;
   title="Target File Type: &filetype" 
     %if %length(&filetype) = 0 %then %do;
      ||' NOT SEPCIFIED';
     %end;
     ;output;
   title="Target File Name Pattern: &filenamePattern" 
     %if %length(&filenamePattern) = 0 %then %do;
       ||' NOT SEPCIFIED';
     %end;
     ;output;
   title="Recursive Search: %upcase(&recursive)"; output;
   title="# of Target Files: &nFiles";output;
   title="# of Files Containg Search String: &nfound";  output;
   title="Generation Date: &rundt"; output;
   title=" "; output;
   title=repeat('#', 139); output;
   %do i = 1 %to 10;
     %if %length(&&searchStr&i) > 0 %then %do;
       title="Search String &i("||"Regular Expression: &&_s_reg&i  "
             ||"Case Sensitive: &&_s_case&i): "||&&_searchStr&i; 
       output;
     %end;
     %if %length(&&excludeStr&i) > 0 %then %do;
       title="Exclude String &i("||"Regular Expression: &&_e_reg&i  "
             ||"Case Sensitive: &&_e_case&i): "||&&_excludeStr&i; 
       output;
     %end;
   %end;
   title=repeat('#', 139); output;
   title=" "; output;
   title=" "; output;
run;

option nodate nonumber ;
title; footnote;
ods _all_ close;
options ps=80 ls=140;
options formchar = "|----|+|---+=|-/\<>*";
options byline center;
ods listing;
proc print data=_title noobs split='~';
  var title;
  format title $256.;
  label title = '~';
run;

%if &nfound>0 %then %do;
  proc sort data=matchFiles; by filewpath linenum; run;
  proc print data=matchFiles split= '~' noobs;
    by filewpath notsorted;
    var linenum / style(column)=[just=l];
    var line;
    format line $120.;
    label line = '~'  linenum='~';
  run;
%end;
%else %do;
  data matchFiles;
    length a $255;
    a='No Results Found';
  run;

  proc print data=matchFiles noobs split='~';
    var a;
    label a = '~';
  run;
%end;
ods listing close;

%mend strSearch;

