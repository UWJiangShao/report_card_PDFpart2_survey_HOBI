OPTIONS PS=MAX FORMCHAR="|----|+|---+=|-/\<>*" MLOGIC MPRINT SYMBOLGEN noxwait noxsync;

LIBNAME IN02 "C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\5. Composite\Data\raw_data\survey";
%LET JOB = P32;


*** ------ STAR Child prepare data ----------------------------------------------------------------------------------------------------------;
proc import datafile="C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\3. Survey\Data\Temp_Data\SC24_out_rate.xlsx"
	dbms=XLSX
	out=SC24
;
	sheet="SC24_rate";
run;


*** ------ STAR Adult prepare data ----------------------------------------------------------------------------------------------------------;
proc import datafile="C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\3. Survey\Data\Temp_Data\SA24_out_rate.xlsx"
	dbms=XLSX
	out=SA24
;
	sheet="SA24_rate";
run;

*** ------ STAR + PLUS prepare data ----------------------------------------------------------------------------------------------------------;
proc import datafile="C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\3. Survey\Data\Temp_Data\SP24_out_rate.xlsx"
	dbms=XLSX
	out=SP24
;
	sheet="SP24_rate";
run;

*** ------ STAR Kids prepare data ----------------------------------------------------------------------------------------------------------;
proc import datafile="C:\Users\jiang.shao\Dropbox (UFL)\MCO Report Card - 2024\Program\3. Survey\Data\Temp_Data\SK24_out_rate.xlsx"
	dbms=XLSX
	out=SK24
;
	sheet="SK24_rate";
run;

proc contents data= SK24 varnum;
run;

/* Prepare for merged dataset */ 
data IN02.SC_survey; 
	set SC24;
	rename PHI_Plan_Code = plancode;
	keep PHI_Plan_Code GCQ_rat--HPRat_rat;
run;

data IN02.SA_survey;
	set SA24;
	rename PHI_Plan_Code = plancode;
	keep PHI_Plan_Code AtC_rat--HPRat_rat;
run;

data IN02.SP_survey;
	set SP24;
	rename PHI_Plan_Code = plancode;
	keep PHI_Plan_Code ATC_rat--HPRat_rat;
run;

data IN02.SK_survey;
	set SK24;
	rename PHI_Plan_Code = plancode;
	rename APM_rat = APM_survey_rat;
	keep PHI_Plan_Code HPRat_rat--BHCoun_rat;
run;