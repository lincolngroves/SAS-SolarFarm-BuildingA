*---------------------------------------------------------------------------------------*
|       	     	       Data & Analytics for Good Journal							|
|  				00 - SAS Building A Consumption Analysis - 09.26.2023 					|
*---------------------------------------------------------------------------------------*;
libname 	solar "C:\Users\ligrov\OneDrive - SAS\Data and Analytics for Good Journal\Solar Farm\SAS Data";
options 	orientation=landscape mlogic symbolgen pageno=1 error=3;

title1 		"Data & Analytics for Good Journal";
title2 		"SAS Building A Consumption Analysis";
footnote 	"File = 00 - SAS Building A Consumption Analysis - 09.26.2023";


*-----------------------------------------------------------------------------------------------*
|                    		Import Raw Data from SAS Solar Farm + Building A        			|      
*-----------------------------------------------------------------------------------------------*;
proc import datafile="C:\Users\ligrov\OneDrive - SAS\Data and Analytics for Good Journal\Solar Farm\Raw Data\BldgASF1-15min_Lincoln_Received on July 28, 2023.xlsx"
     out=solar_15 
     dbms=xlsx
     replace;
     getnames=yes;
	 sheet="Data";
run;


*-----------------------------------------------------------------------------------------------*
|                    		Collapse Data + Create RHS Modeling Variables		       			|      
*-----------------------------------------------------------------------------------------------*;
proc sql;
	create 	table solar.solar_daily as
	select	distinct date,
			count( * ) as Sanity_Check,
			sum( Bldg_A_Consumption_from_Utility ) 	as Bldg_A_Consumption 		label="Bldg A Consumption from Utility (kWh)"	format comma11. ,
			log( calculated Bldg_A_Consumption ) 	as Log_Bldg_A_Consumption	label="Log of Bldg A Consumption (kWh)"			format 9.4 ,
			sum( SF1_kWh_Generated ) 				as Solar_Farm_Generation	label="Solar Farm 1 | kWh Generated"			format comma11. ,
			( weekday(date)=1 or weekday(date)=7 )  as Weekend					label="Weekend"									format 4.,
			month(date)								as Month					label="Month"									format 4.,
			case 	when 				 date <  "01APR2020"d 	then	4 	
					when "01APR2020"d <= date <  "01JAN2022"d 	then	1
					when "01JAN2022"d <= date <  "01FEB2023"d	then 	2
					when 				 date >= "01FEB2023"d	then 	3
																else	.
			end		as COVID_Period		label="COVID Period"
	from 	solar_15
	group	by 1 
	having	sanity_check = 96;
quit;


*-----------------------------------------------------------------------------------------------*
|								Regression Modeling | Actual Consumption       					|      
*-----------------------------------------------------------------------------------------------*;
ods listing close;
ods pdf file="C:\Users\ligrov\OneDrive - SAS\Data and Analytics for Good Journal\Solar Farm\Output\SAS Building A Consumption Regressions - &sysdate..pdf";


	*****************************************  All Days Pre-Post ;
	proc glm data=solar.solar_daily;
	   class 	weekend month covid_period; 
	   model 	Bldg_A_Consumption = covid_period month / solution; 
	   ods 		output ParameterEstimates=parameter_estimates1 ;
	run; quit;


	*****************************************  Relative to Weekend ;
	proc glm data=solar.solar_daily;
	   class 	weekend month covid_period; 
	   model 	Bldg_A_Consumption = weekend covid_period weekend*covid_period month / solution; 
	   ods 		output ParameterEstimates=parameter_estimates2 ;
	run; quit;


	*-----------------------------------------------------------------------------------------------*
	|								Regression Modeling | Log Consumption       					|      
	*-----------------------------------------------------------------------------------------------*;

	*****************************************  All Days Pre-Post ;
	proc glm data=solar.solar_daily;
	   class 	weekend month covid_period; 
	   model 	Log_Bldg_A_Consumption = covid_period month / solution; 
	   ods 		output ParameterEstimates=parameter_estimates3 ;
	run; quit;


	*****************************************  Relative to Weekend ;
	proc glm data=solar.solar_daily;
	   class 	weekend month covid_period; 
	   model 	Log_Bldg_A_Consumption = weekend covid_period weekend*covid_period month / solution; 
	   ods 		output ParameterEstimates=parameter_estimates4 ;
	run; quit;


ods pdf close;
ods listing;


*-----------------------------------------------------------------------------------------------*
|								Export Parameter Estimates Tables to Excel     					|      
*-----------------------------------------------------------------------------------------------*;

%macro hilfe();
	%do i=1 %to 4;
		proc export data=parameter_estimates&i
			outfile="C:\Users\ligrov\OneDrive - SAS\Data and Analytics for Good Journal\Solar Farm\Output\Parameter Estimates - &sysdate..xlsx"
			dbms=xlsx replace;
			sheet="parameter_estimates&i";
		run;
	%end;
%mend;

%hilfe();
