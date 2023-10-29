%macro execUnixCommand(command);

    /* Define a filename for the PIPE option */
    filename cmdout pipe "&command.";

    /* Read command output directly to the SAS log */
    data _null_;
        infile cmdout truncover;
        input line $char200.;
        put line; /* Write each line to the SAS log */
    run;

    filename cmdout clear;

    /* Check the return code and display a message in the SAS log */
    %if &SYSCC. = 0 %then %do;
        %put NOTE: The Unix command executed successfully.;
    %end;
    %else %do;
        %put ERROR: The Unix command failed with return code &SYSCC..;
    %end;

%mend execUnixCommand;

/* Usage example */
%execUnixCommand(ls -l);
