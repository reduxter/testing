/* Define your SAS datasets */
%let email_subject = "Your Email Subject";

/* Define specific text for each dataset */
%let text_dataset1 = "This is specific text for YourDataset1.";
%let text_dataset2 = "This is specific text for YourDataset2.";
%let text_dataset3 = "This is specific text for YourDataset3.";
%let text_dataset4 = "This is specific text for YourDataset4.";
%let text_dataset5 = "This is specific text for YourDataset5.";
%let text_dataset6 = "This is specific text for YourDataset6.";

/* Define network drive path */
%let network_drive_path = "Your network drive path";

/* Start ODS HTML output */
ods html body="your_email_body.html" style=sasweb;

/* Begin email body */
ods escapechar='^';
ods _all close;
ods listing close;
ods proclabel="Your Email Body";
title "Your Email Title";

/* Check for DIFF_ columns */
%macro check_diff_columns(dataset);
    proc sql noprint;
        select 
            case when count(*) > 0 then 1 else 0 end as has_diff,
            name into :has_diff, :diff_col
        from dictionary.columns
        where libname='WORK' and memname=upcase("&dataset") and name like 'DIFF_%';
    quit;
%mend;

%check_diff_columns(YourDataset1);
%check_diff_columns(YourDataset2);
%check_diff_columns(YourDataset3);
%check_diff_columns(YourDataset4);
%check_diff_columns(YourDataset5);
%check_diff_columns(YourDataset6);

/* Generate HTML content */
data _null_;
    file print;
    put "<html>";
    put "<body>";

    /* Start email body */
    put "<p>Hello,</p>";
    put "<br>"; /* blank line */
    put "<p>This is the result of the comparison:</p>";

    /* Process dataset YourDataset1 */
    put "<p>&text_dataset1</p>";
    put "<p>Dataset YourDataset1:</p>";
    %if &has_diff %then %do;
        put "<p>YourDataset1 has columns starting with 'DIFF_'. Here are the columns:</p>";
        put "<ul><li>&diff_col</li></ul>";
    %end;
    %else %do;
        put "<p>YourDataset1 does not have columns starting with 'DIFF_'.</p>";
    %end;
    put "<p>Here is the dataset:</p>";
run;
proc print data=YourDataset1 noobs; run;

data _null_;
    file print;
    put "<br>"; /* blank line */

    /* Process dataset YourDataset2 */
    put "<p>&text_dataset2</p>";
    put "<p>Dataset YourDataset2:</p>";
    %if &has_diff %then %do;
        put "<p>YourDataset2 has columns starting with 'DIFF_'. Here are the columns:</p>";
        put "<ul><li>&diff_col</li></ul>";
    %end;
    %else %do;
        put "<p>YourDataset2 does not have columns starting with 'DIFF_'.</p>";
    %end;
    put "<p>Here is the dataset:</p>";
run;
proc print data=YourDataset2 noobs; run;

data _null_;
    file print;
    put "<br>"; /* blank line */

    /* Process dataset YourDataset3 */
    put "<p>&text_dataset3</p>";
    put "<p>Dataset YourDataset3:</p>";
    %if &has_diff %then %do;
        put "<p>YourDataset3 has columns starting with 'DIFF_'. Here are the columns:</p>";
        put "<ul><li>&diff_col</li></ul>";
    %end;
    %else %do;
        put "<p>YourDataset3 does not have columns starting with 'DIFF_'.</p>";
    %end;
    put "<p>Here is the dataset:</p>";
run;
proc print data=YourDataset3 noobs; run;

data _null_;
    file print;
    put "<br>"; /* blank line */

    /* Process dataset YourDataset4 */
    put "<p>&text_dataset4</p>";
    put "<p>Dataset YourDataset4:</p>";
    %if &has_diff %then %do;
        put "<p>YourDataset4 has columns starting with 'DIFF_'. Here are the columns:</p>";
        put "<ul><li>&diff_col</li></ul>";
    %end;
    %else %do;
        put "<p>YourDataset4 does not have columns starting with 'DIFF_'.</p>";
    %end;
    put "<p>Here is the dataset:</p>";
run;
proc print data=YourDataset4 noobs; run;

data _null_;
    file print;
    put "<br>"; /* blank line */

    /* Process dataset YourDataset5 */
    put "<p>&text_dataset5</p>";
    put "<p>Dataset YourDataset5:</p>";
    %if &has_diff %then %do;
        put "<p>YourDataset5 has columns starting with 'DIFF_'. Here are the columns:</p>";
        put "<ul><li>&diff_col</li></ul>";
    %end;
    %else %do;
        put "<p>YourDataset5 does not have columns starting with 'DIFF_'.</p>";
    %end;
    put "<p>Here is the dataset:</p>";
run;
proc print data=YourDataset5 noobs; run;

data _null_;
    file print;
    put "<br>"; /* blank line */

    /* Process dataset YourDataset6 */
    put "<p>&text_dataset6</p>";
    put "<p>Dataset YourDataset6:</p>";
    %if &has_diff %then %do;
        put "<p>YourDataset6 has columns starting with 'DIFF_'. Here are the columns:</p>";
        put "<ul><li>&diff_col</li></ul>";
    %end;
    %else %do;
        put "<p>YourDataset6 does not have columns starting with 'DIFF_'.</p>";
    %end;
    put "<p>Here is the dataset:</p>";
run;
proc print data=YourDataset6 noobs; run;

data _null_;
    file print;
    put "<br>"; /* blank line */

    /* Additional datasets */
    put "<p>Additional Datasets:</p>";
    put "<br>"; /* blank line */
    put "<p>Dataset YourDataset7:</p>";
run;
proc print data=YourDataset7 noobs; run;

data _null_;
    file print;
    put "<br>"; /* blank line */
    put "<p>Dataset YourDataset8:</p>";
run;
proc print data=YourDataset8 noobs; run;

data _null_;
    file print;
    put "<br>"; /* blank line */

    /* Information about results saved on an Excel sheet */
    put "<p>The results are also saved on an Excel sheet located on the network drive.</p>";
    put "<p>Network Drive Path:</p>";
    put "<p><a href='&network_drive_path'>&network_drive_path</a></p>";

    /* Please review and Thank you */
    put "<p>Please review.</p>";
    put "<p>Thank you!</p>";
    put "</body>";
    put "</html>";
run;

/* End HTML email body */
ods html close;

/* Send the email */
filename mymail email
    to="recipient@example.com"
    subject="&email_subject"
    type="text/html"
    attach="your_email_body.html";

data _null_;
    file mymail;
    /* Include the HTML content of the email body */
    infile "your_email_body.html" lrecl=32767;
    input;
    put _infile_;
run;

/* Close the email file */
filename mymail clear;
