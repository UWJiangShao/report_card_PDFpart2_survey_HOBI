OPTIONS PS=MAX FORMCHAR="|----|+|---+=|-/\<>*" MLOGIC MPRINT SYMBOLGEN;

* Prepare the dataframe for survey rating;

%LET JOB = P12;

LIBNAME IN01 "..\DATA\Raw_Data\";
LIBNAME OUT "..\DATA\Temp_Data\";

* macro to import all sheets from one Excel file;
%macro import_excel(excel_address, prog);

	libname ELIB excel &excel_address. mixed=yes;

	proc sql;
		create table excel_sheet as
		select *
		from dictionary.tables
		where libname = "ELIB"
		;

		select memname
		into :sheet_list separated by '*'
		from excel_sheet
		;

		select count(memname)
		into :sheet_n
		from excel_sheet
		;
	quit;

	%put &sheet_list.;
	%put &sheet_n.;

	%macro import_sheet;
		%do i = 1 %to &sheet_n.;
			%let var = %scan(&sheet_list., &i., *);
			%if %sysfunc(find(&var., $)) = 0 %then %do;
				%let sf = %substr(&var, 1, %length(&var.));
				data &prog._&sf.;
					set ELIB."&var."n;
					where PHI_Plan_Code ne ' ';
				run;

				proc contents data=&prog._&sf. varnum;
				run;

				proc sort data=&prog._&sf.;
					by PHI_Plan_Code;
				run;
			%end;
		%end;
	%mend import_sheet;

	%import_sheet;

%mend import_excel;


%macro import_pop(prog);
	data &prog._population;
		set IN01.P03_df_&prog._weight;
		rename
			sample_pool_members = Population
			;
	run;

	proc sort data=&prog._population;
		by PHI_Plan_Code;
	run;

	proc contents data=&prog._population varnum;
	run;
%mend import_pop;

/* Macro to delete excel for fully replacement */
%MACRO XLSX_BAK_DELETE(file);
	DATA _NULL_;
		FNAME = 'TODELETE';
		RC = FILENAME(FNAME, &file.);
		RC = FDELETE(FNAME);
		RC = FILENAME(FNAME);
	RUN;
%MEND XLSX_BAK_DELETE;


* For STAR Child ----------------------------------------------------------------------------------------;
%import_excel("..\DATA\Raw_Data\P03c_df_SC_merge_w_out_V1.xlsx", SC);
%import_pop(SC);

data STAR_Child_for_analysis;
	merge 
		SC_GCQ
		SC_HPRAT
		SC_HWDC
		SC_PDRAT
		SC_ROUTCARE
		SC_URGCARE
		SC_population(keep=PHI_Plan_Code Population)
		;
	by PHI_Plan_Code;
run;


* Delete the existing file ;
%let SC_out_xlsx = '..\DATA\Temp_Data\SC24_out.xlsx';
%XLSX_BAK_DELETE(&SC_out_xlsx.);

proc export data=STAR_Child_for_analysis
	outfile=&SC_out_xlsx.
	dbms=xlsx
	replace
	;
run;


* For STAR Adults ---------------------------------------------------------------------------------------;
%import_excel("..\DATA\Raw_Data\P03_df_SA_merge_w_out_V1.xlsx", SA);
%import_pop(SA);

data STAR_Adult_for_analysis;
	merge 
		SA_ATC
		SA_GCQ
		SA_GNC
		SA_HPRAT
		SA_HWDC
		SA_PDRAT
		SA_population(keep=PHI_Plan_Code Population)
		;
	by PHI_Plan_Code;
run;

* Delete the existing file ;
%let SA_out_xlsx = '..\DATA\Temp_Data\SA24_out.xlsx';
%XLSX_BAK_DELETE(&SA_out_xlsx.);

proc export data=STAR_Adult_for_analysis
	outfile=&SA_out_xlsx.
	dbms=xlsx
	replace
	;
run;


* For STAR PLUS -----------------------------------------------------------------------------------------;
%import_excel("..\DATA\Raw_Data\P03_df_SP_merge_w_out_V1.xlsx", SP);
%import_pop(SP);


DATA STARPLUS_for_analysis;
	merge 
		SP_ATC
		SP_GCQ
		SP_GNC
		SP_HPRAT
		SP_HWDC
		SP_PDRAT
		SP_population(keep=PHI_Plan_Code Population)
		;
	by PHI_Plan_Code;
run;

* Delete the existing file ;
%let SP_out_xlsx = '..\DATA\Temp_Data\SP24_out.xlsx';
%XLSX_BAK_DELETE(&SP_out_xlsx.);

proc export data=STARPLUS_for_analysis
	outfile=&SP_out_xlsx.
	dbms=xlsx
	replace
	;
run;


* For STAR Kids -----------------------------------------------------------------------------------------;
%import_excel("..\DATA\Raw_Data\P03_df_SK_merge_w_out_V1.xlsx", SK);
%import_pop(SK);

DATA STARKIDS_for_analysis;
	MERGE 
		SK_ATC
		SK_GCQ
		SK_GNC
		SK_HPRAT
		SK_SPECTHER
		SK_COORD
		SK_BHCOUN
		SK_CCCGNI
		SK_CCCMEDS
		SK_TRTADULT
		SK_population(keep=PHI_Plan_Code Population)
		;
	by PHI_Plan_Code;
run;

data STARKIDS_for_analysis;
	set STARKIDS_for_analysis;
	rename 
		CCCGNI = GNI
		CCCGNI_den = GNI_den
		CCCGNI_stderr = GNI_stderr
		CCCGNI_sig = GNI_sig
		CCCMeds = APM
		CCCMeds_den = APM_den
		CCCMeds_stderr = APM_stderr
		CCCMeds_sig = APM_sig
		trtadult = transit
		trtadult_den = transit_den
		trtadult_stderr = transit_stderr
		trtadult_sig = transit_sig
		;
run;

* Delete the existing file ;
%let SK_out_xlsx = '..\DATA\Temp_Data\SK24_out.xlsx';
%XLSX_BAK_DELETE(&SK_out_xlsx.);

proc export data=STARKIDS_for_analysis
	outfile=&SK_out_xlsx.
	dbms=xlsx
	replace
	;
run;
