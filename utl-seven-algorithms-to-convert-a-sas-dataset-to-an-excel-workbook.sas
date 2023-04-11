%let pgm=utl-seven-algorithms-to-convert-a-sas-dataset-to-an-excel-workbook;

Seven algorithms to convert a sas dataset to an excel workbook


How to get VBA code with SAS Addin to work (InsertDataFromSASFolder)

Highligt a SAS data or copy sas dataset name to the clipboard and view the dataset in excel

   1. Colectica Product ($50 perpetual) addin to excel. Do not need SAS or WPS)
   2. SAS Classic Editor command xlrh  (uses proc report)
   3. SAS Classic editor command xlsh (uses libname - could passthru to ant database))
   4. WPS (uses proc report)
   5. WPS (uses libname - could passthru to ant database))
   6. R
   7. Python


github
https://tinyurl.com/yc678sfe
https://github.com/rogerjdeangelis/utl-seven-algorithms-to-convert-a-sas-dataset-to-an-excel-workbook

Related
https://tinyurl.com/4e33ku73
https://stackoverflow.com/questions/75439184/how-to-get-vba-code-with-sas-addin-to-work-insertdatafromsasfolder

/*      ____      _           _   _
/ |    / ___|___ | | ___  ___| |_(_) ___ __ _
| |   | |   / _ \| |/ _ \/ __| __| |/ __/ _` |
| |_  | |__| (_) | |  __/ (__| |_| | (_| (_| |
|_(_)  \____\___/|_|\___|\___|\__|_|\___\__,_|

*/

Colectica
https://www.colectica.com/software/colecticaforexcel/

$50

All Standard Features
Perpetual License
Import Spss to Excel
Import Stata to Excel
Import SAS to Excel
Email Support

/*___                           _      _
|___ \     ___  __ _ ___  __  _| |_ __| |__
  __) |   / __|/ _` / __| \ \/ / | `__| `_ \
 / __/ _  \__ \ (_| \__ \  >  <| | |  | | | |
|_____(_) |___/\__,_|___/ /_/\_\_|_|  |_| |_|

*/

%macro xlrh /cmd des="Usage: xlrh. Hilite a table and type xlrh and table will open it in excel. No need for pc acces to excel";
   store;note;notesubmit '%xlrha;';
   run;
%mend xlrh;

%macro xlrha/cmd;

    %local argx;

    filename clp clipbrd ;

    data _null_;
       infile clp;
       input;
       argx=_infile_;
       call symputx("argx",argx);
       putlog argx=;
    run;quit;

    /* %let argx=sashelp.class; */

    ods escapechar = '~';

    %utlfkil(%sysfunc(getoption(work))/_rpt.xlsx);

    ods listing close;

    ods escapechar='~';

    %utl_xlslan100;

    ods excel file="%sysfunc(getoption(work))/_rpt.xlsx"
            options(
              /* autofit_height           = 'yes'*/
               sheet_name                 = "&argx"
               autofilter                 = "yes"
               frozen_headers             = "1"
               gridlines                  = "yes"
               embedded_titles            = "yes"
               embedded_footnoteS         = "yes"
               );

    proc report data=&argx missing
     style(column)=[textalign=left verticalalign=top cellwidth=8in] split="~" style=utl_xlslan100;
    title "SAS table &argx";
    run;quit;

    ods excel close;

    ods listing;

    options noxwait noxsync;
    /* Open Excel */
    x "excel %sysfunc(getoption(work))/_rpt.xlsx";
    run;quit;

%mend xlrha;

/*____                         _     _
|___ /    ___  __ _ ___  __  _| |___| |__
  |_ \   / __|/ _` / __| \ \/ / / __| `_ \
 ___) |  \__ \ (_| \__ \  >  <| \__ \ | | |
|____(_) |___/\__,_|___/ /_/\_\_|___/_| |_|

*/

%macro xlsh /cmd des="Usage: xlsh. Highlight table and type xlsh and the table  will appear in excel uses libname need pc acces";;
   store;note;notesubmit '%xlsha;';
   run;
%mend xlsh;

%macro xlsha/cmd;

    filename clp clipbrd ;
    data _null_;
     infile clp;
     input;
     put _infile_;
     call symputx('argx',_infile_);
    run;

    %let __tmp=%sysfunc(pathname(work))\myxls.xlsx;

    data _null_;
        fname="tempfile";
        rc=filename(fname, "&__tmp");
        put rc=;
        if rc = 0 and fexist(fname) then
       rc=fdelete(fname);
    rc=filename(fname);
    run;

    libname __xls excel "&__tmp";
    data __xls.%scan(__&argx,1,%str(.));
        set &argx.;
    run;quit;
    libname __xls clear;

    data _null_;z=sleep(1);run;quit;

    options noxwait noxsync;
    /* Open Excel */
    x "excel &__rpt.xlsx";
    run;quit;

%mend xlsha;

/*  _                                _      _
| || |    __      ___ __  ___  __  _| |_ __| |__
| || |_   \ \ /\ / / `_ \/ __| \ \/ / | `__| `_ \
|__   _|   \ V  V /| |_) \__ \  >  <| | |  | | | |
   |_|(_)   \ _/\_/| .__/|___/ /_/\_\_|_|  |_| |_|
                    |_|
*/

/*---- you need to highlight class and store in the name in the clipboard (ctrl c) ----*/;
/*---- class

%let _wrk=%sysfunc(pathname(work));
%put &=_wrk;

data class;
  set sashelp.class;
run;quit;

%utl_submit_wps64("

    libname wrk '&_wrk';

    filename clp clipbrd ;

    data _null_;
       infile clp;
       input;
       argx=_infile_;
       call symputx('argx',argx);
       putlog argx=;
    run;quit;

    /*---- %let argx=class; ----*/

    ods escapechar = '~';

    ods listing close;

    ods escapechar='~';

    ods excel file='%sysfunc(getoption(work))/_rpt.xlsx'
            options(
              /* autofit_height           = 'yes'*/
               sheet_name                 = '&argx'
               frozen_headers             = '1'
               gridlines                  = 'yes'
               embedded_titles            = 'yes'
               embedded_footnoteS         = 'yes'
               );

    proc report data=wrk.&argx missing
     style(column)=[textalign=left verticalalign=top cellwidth=8in] split='~' style=utl_xlslan100;
    title 'SAS table &argx';
    run;quit;

    ods excel close;

    ods listing;

    options noxwait noxsync;
    /* Open Excel */
    x 'excel %sysfunc(getoption(work))/_rpt.xlsx';
    run ;quit;
");

/*___                               _     _
| ___|   __      ___ __  ___  __  _| |___| |__
|___ \   \ \ /\ / / `_ \/ __| \ \/ / / __| `_ \
 ___) |   \ V  V /| |_) \__ \  >  <| \__ \ | | |
|____(_)   \_/\_/ | .__/|___/ /_/\_\_|___/_| |_|
                  |_|
*/

/*---- you need to highlight class and store in the name in the clipboard (ctrl c) ----*/;
/*---- class

%let _wrk=%sysfunc(pathname(work));
%put &=_wrk;

data class;
  set sashelp.class;
run;quit;

%utl_submit_wps64("

    libname wrk '&_wrk';

    /*---- read the clipboard witht he dataset name you loaded into the clipboard previously ----*/
    filename clp clipbrd ;
    data _null_;
     infile clp;
     input;
     put _infile_;
     call symputx('argx',_infile_);
    run;

    %let __tmp=%sysfunc(pathname(work))\myxls.xlsx;

    data _null_;
        fname='tempfile';
        rc=filename(fname, '&__tmp');
        put rc=;
        if rc = 0 and fexist(fname) then
       rc=fdelete(fname);
    rc=filename(fname);
    run;

    libname __xls excel '&__tmp';
    data __xls.%scan(__&argx,1,%str(.));
        set wrk.&argx.;
    run;quit;
    libname __xls clear;

    data _null_;z=sleep(1);run;quit;

    options noxwait noxsync;
    /* Open Excel */
    x 'excel &__tmp';
    run;quit;
");

/*__      ____
 / /_    |  _ \
| `_ \   | |_) |
| (_) |  |  _ <
 \___(_) |_| \_\

*/

libname sd1 "d:/sd1";

data sd1.class;
  set sashelp.class;
run;quit;

%utl_submit_r64("
library(haven);
library(XLConnect);
class<-read_sas('d:/sd1/class.sas7bdat');
class;
wb <- loadWorkbook('d:/xls/xis.xlsx',create=TRUE);
createSheet(wb, name = 'CLASS');
writeWorksheet(wb, class, sheet = 'CLASS');
saveWorkbook(wb);
");

options noxwait noxsync;
/* Open Excel */
x 'excel d:/xls/xis.xlsx';
run ;quit;

/*____               _   _
|___  |  _ __  _   _| |_| |__   ___  _ __
   / /  | `_ \| | | | __| `_ \ / _ \| `_ \
  / /_  | |_) | |_| | |_| | | | (_) | | | |
 /_/(_) | .__/ \__, |\__|_| |_|\___/|_| |_|
        |_|    |___/
*/

/*---- could use pyreadstat ----*/

libname sd1 "d:/sd1";

data sd1.class;
  set sashelp.class;
run;quit;

%utl_submit_py64_310("
import pandas as pd;
from sas7bdat import SAS7BDAT;
with SAS7BDAT('d:/sd1/class.sas7bdat') as m:;
.   clas = m.to_data_frame();
clas.to_excel(r'd:\xls\clas.xlsx', index=False) ;
");

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
