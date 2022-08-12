/*************************************************************************************************
Program Name    : RiskEst.sas
Purpose         : Estimation of risk parameters and make adjustment in case of empty cell 
Project         : GitHub Archive
Programmer      : Cha
Version/Date    : 1.0 02JUN2022
SAS Version     : 9.4
Comments        :
Modifications   : 
**************************************************************************************************/


/***************************************************************/
/*  Run the macro once, then you can use the code in the last **/
/* line to calculate everything you need************************/
/***************************************************************/

%let path=C:\Users\cengiz.acikel\Desktop\Cengiz\ArbeitDokumente\SoftwareDocuments\SAS_documents\MyGitHub\RiskEstimationWithNullCell;

%macro riskest (acresp, acnoresp, plcresp, plcnorsp);

DATA Design;
DO Arm = "Act","Plc";
    DO Avalc = "Resp","NoRe", "Ttl";
        output;
    	END;
	END;
RUN;

Data Design;
Set Design;
if Arm="Act" & Avalc="Resp" then value=&acresp;
if Arm="Act" & Avalc="NoRe" then value=&acnoresp;
if Arm="Act" & Avalc="Ttl" then value=&acresp+&acnoresp;
if Arm="Plc" & Avalc="Resp" then value=&plcresp;
if Arm="Plc" & Avalc="NoRe" then value=&plcnorsp;
if Arm="Plc" & Avalc="Ttl" then value=&plcnorsp+&plcresp;
format avalc $15.;
run;

proc sql;
create table design1 as(select*,min(value) as mn_value from design);
run;

Data Design1;
Set Design1;
if mn_value>0  then do;
weight=value;
end;
else weight=value+0.5;
run;

Proc Freq data=Design1;
where avalc="Resp" OR avalc="NoRe";
table arm*Avalc/NOROW NOPERCENT	NOCUM CHISQ	RISKDIFF RELRISK SCORES=TABLE ALPHA=0.05	OUT=TableDes(LABEL="Zellenstatistiken für Avalc nach Arm");
	OUTPUT 	CHISQ RISKDIFF 	RELRISK OUT=TblStats(LABEL="Tabellenstatistiken für Avalc nach Arm");
weight weight;
run;

proc sort data=Design1;
by Arm;
run;

proc transpose data=Design1 out=CrossTab (drop= _name_ rename=(col1=Resp col2=NoRe col3=Ttl)) ;
var value;
by arm;
run;

Data CrsTabP;
set CrossTab;
RespP=(Resp/(Resp+NoRe))*100;
NoReP=(NoRe/(Resp+NoRe))*100;
format RespP NoReP 4.1;
RespVal=Resp||" ("||put(RespP,4.1)||"%)";
NoReVal=NoRe||" ("|| put (NoReP,4.1)||"%)";
TtlVal=ttl||" (100%)";
keep arm RespVal NoReVal TtlVal;
run;

proc transpose data=CrsTabP out=CrsTabPH;
var RespVal NoReVal TtlVal;
by Arm;
run;

proc sort data= Crstabph;
by _name_;
run;

proc transpose data=Crstabph out=Crstabpf (rename=(col1=Trt col2=Plcb)) ;
var col1;
by _name_;
run;

proc transpose data=Design out=Fisher;
var value;
run;

data TblStats;
set TblStats;
mergevar="001";
run;

data Fisher;
set Fisher;
mergevar="001";
run;

data TblStaMr;
merge TblStats Fisher;
by mergevar;
run;

Proc Format;
value Parametr
0 = "Odd's ratio (OR)"                           
1 = "Attributable risk ratio (ARR)"                              
2 = "Relative risk (RR)" 
3 = "Number Needed to Treat(NNT)"                             
4 = "p value";
value $avalc
"NoReVal"="Not respondent"
"RespVal"="Respondent"
"TtlVal"="Total";
run;

   data Stats;
   set TblStaMr;
   Parametr = 0;
   Value = put(_RROR_, 8.3) ;
   confint = put(L_RROR, 8.3)||" - "||put(U_RROR, 8.3);
   output;
   Parametr = 1;
   Value = put(_RDIF1_,8.3);
   confint = put(L_RDIF1, 8.3)||" - "||put(U_RDIF1, 8.3);
   output;
   Parametr = 2;
   Value = put(_RRC1_,8.3);
   confint = put(L_RRC1, 8.3)||" - "||put(U_RRC1, 8.3);
   output;
   Parametr = 3;
   Value = put(1/_RDIF1_,8.3);
   confint = put(1/U_RDIF1, 8.3)||" - "||put(1/L_RDIF1, 8.3);
   output;
   Parametr = 4;
   if ((col1<5 and col1 gt 0) or (col2<5 and col2 gt 0) or col4<5  and (col4 gt 0)or (col4<5 and col4 gt 0)) then do value=put(XP2_FISH,8.3);
   confint="Fisher's exact test";
   end;
   else if (col1 GE 5 AND col2 GE 5 AND col4 GE 5 AND col5 GE 5) OR (col1=0 OR col2=0 OR col4=0 OR col5=0)then do value=put(P_PCHI,8.3);
   confint="Chi-square test";
   end;
   else if (col1 LT 0 OR col2 LT 0 OR col4 LT 0 OR col5 LT 0) then do value="";
   confint="NA (Please check cell contents";
   end;
   output;
   keep Parametr Value Confint;
   format Parametr Parametr.;
run;

Data Final;
merge Crstabpf Stats;
rownum=_n_;
format _name_ avalc.
run;

ods rtf file = "&path\RiskEst.rtf";
ods escapechar="^";
options nodate nonumber orientation=landscape; 
proc report data=Final nowd split='*' missing wrap
style(report)={just=center rules=none frame=hsides cellspacing=0 cellpadding=0}
style(lines)=header{background=white asis=on font_size=12pt font_face="TimesRoman"
font_weight=bold just=left}
style(header)=header{background=cxD9D9D9 font_size=10pt font_face="TimesRoman" frame=box
font_weight=bold}
style(column)=header{background=white font_size=10pt font_face="TimesRoman"
font_weight=medium};
columns  rownum _name_  Trt Plcb Parametr Value Confint;
define rownum / 'No' order order=data width=6 style(column)={just=c};
define _name_ / 'Study arm' group order style(column)={just=l};
define Trt / 'Number and percentage of subjects in treatment group' group order style(column)={just=c};
define Plcb / 'Number and percentage of subjects in placebo group' group order style(column)={just=c};
define Parametr / 'Parameters ^{super *, 1-3}' order order=data width=20 style(column)={just=left};
define Value / 'Effect size estimations' order order=data width=20 style(column)={just=c};
define confint / '95% CI estimations' order order=data width=20 style(column)={just=c};
compute after / style={just=left};
line '^{super *}^S={width=100% just=c font_face=TimesRoman fontsize=10pt}If there are problems with zero cells when calculating the parameters and their confidence intervals, 0.5 is added to all cells, as defined by Pagano.';
line '^{super 1}^S={width=100% just=c font_face=TimesRoman fontsize=10pt}Pagano M, Gauvreau K, Mattie H. Principles of biostatistics: CRC Press; 2022.';
line '^{super 2}^S={width=100% just=c font_face=TimesRoman fontsize=10pt}Altman DG. Practical statistics for medical research Chapman and Hall. London and New York. 1991.';
line '^{super 3}^S={width=100% just=c font_face=TimesRoman fontsize=10pt}Daly LE. Confidence limits made easy: interval estimation using a substitution method. Am J Epidemiol. 1998;147(8):783-90.';
endcomp;
title1 j=c 'Clinical epidemiology measures';
title2 j=c 'Table 1: Summary table';
title3;
title4;
footnote j=l "Table run:  '&sysdate9 &systime'" j=c "Page ^{thispage} of ^{lastpage}" j=r "SAS templates"; 
footnote4 j=l "Output: &path\RiskEst.rtf";
footnote5 j=l ;
run;

ods rtf close;

%mend riskest;
options mprint mlogic;
%riskest (10,15,20,4);

