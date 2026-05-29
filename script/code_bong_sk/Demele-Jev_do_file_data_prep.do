/************************************************************** 

	Project title: Demele-Jev (ANRS0629) - Cambodia  
	Created by: Kennarey Seang
	Created on: Dec 13 2025
	Purpose: Data preparation & monitoring 

**************************************************************/

/*
 Append the screening forms for cohort 1 and 2 
*/ 

clear 

import excel "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/D1_Screening_Form_Cohort1_Baseline.xlsx", sheet("Sheet 1 - D1_Screening_Form_Coh") cellrange(A2:AV204) firstrow case(lower)

replace a1="DEM1-PP-0005" if a1manual=="DEM1-PP-0005"
replace a1="DEM1-PP-0003" if a1manual=="DEM1-PP-0003"
replace a1="DEM1-PP-0002" if a1manual=="DEM1-PP-0002"
replace a1="DEM1-PP-9001" if a1manual=="DEM1-PP-9001"

save "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-screening-cohort1.dta", replace

clear 

import excel "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/D4_Screening_Form_Cohort2_Baseline.xlsx", sheet("Sheet 1 - D4_Screening_Form_Coh") cellrange(A2:AT38) firstrow case(lower)

tostring a1manual, replace 

save "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-screening-cohort2.dta", replace

append using "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-screening-cohort1.dta", force 

save "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-screening-cohort1&2.dta", replace

/*
 Rename and recoding variables 
*/ 

use "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-screening-cohort1&2.dta", clear 

// Study participants

gen cohort=.
replace cohort=1 if regexm(a1, "(^DEM1)")|regexm(a1manual, "(^DEM1)")
replace cohort=2 if regexm(a1, "(^DEM2)")|regexm(a1manual, "(^DEM2)")

label define cohortft 1"cohort1" 2"cohort2", replace 
label value cohort cohortft 

tab cohort 

// Address
gen address=. 
replace address=1 if a6=="KH-1"|scr_a6=="KH-1"
replace address=2 if a6=="KH-2"|scr_a6=="KH-2"
replace address=3 if a6=="KH-3"|scr_a6=="KH-3"
replace address=4 if a6=="KH-4"|scr_a6=="KH-4"
replace address=5 if a6=="KH-5"|scr_a6=="KH-5"
replace address=6 if a6=="KH-6"|scr_a6=="KH-6"
replace address=7 if a6=="KH-7"|scr_a6=="KH-7"
replace address=8 if a6=="KH-8"|scr_a6=="KH-8"
replace address=9 if a6=="KH-9"|scr_a6=="KH-9"
replace address=10 if a6=="KH-10"|scr_a6=="KH-10"
replace address=11 if a6=="KH-11"|scr_a6=="KH-11"
replace address=12 if a6=="KH-12"|scr_a6=="KH-12"
replace address=13 if a6=="KH-13"|scr_a6=="KH-13"
replace address=14 if a6=="KH-14"|scr_a6=="KH-14"
replace address=15 if a6=="KH-15"|scr_a6=="KH-15"
replace address=16 if a6=="KH-16"|scr_a6=="KH-16"
replace address=17 if a6=="KH-17"|scr_a6=="KH-17"
replace address=18 if a6=="KH-18"|scr_a6=="KH-18"
replace address=19 if a6=="KH-19"|scr_a6=="KH-19"
replace address=20 if a6=="KH-20"|scr_a6=="KH-20"
replace address=21 if a6=="KH-21"|scr_a6=="KH-21"
replace address=22 if a6=="KH-22"|scr_a6=="KH-22"
replace address=23 if a6=="KH-23"|scr_a6=="KH-23"
replace address=24 if a6=="KH-24"|scr_a6=="KH-24"
replace address=25 if a6=="KH-25"|scr_a6=="KH-25"

label define addft 1"Bantey Meanchey" 2"Battambang" 3"K-Cham" 4"K-Chhnang" ///
	5"K-Speu" 6"K-Thom" 7"Kampot" 8"Kandal" 9"KohKong" 10"Kratie" 11"MondulK" ///
	12"PP" 13"P-Vihear" 14"P-Veng" 15"Pursat" 16"RatanakK" 17"SR" 18"P-Siha" ///
	19"S-Treng" 20"S-Rieng" 21"Takeo" 22"OtdarM" 23"Kep" 24"Pailin" 25"T-Khmum", replace
label value address addft 

// Study sites

gen site=.
replace site=1 if scr_a3==1|a3==1
replace site=2 if scr_a3==2|a3==2

label define siteft 1"kantha bopha" 2"jayavarman vii", replace 
label value site siteft 

tab site

// Age group
gen agecat=. 
replace agecat=1 if scr_a4==1|a4==1
replace agecat=2 if scr_a4==2|a4==2
replace agecat=3 if scr_a4==3|a4==3

label define agecatft 1"2-5 yrs" 2"6-10 yrs" 3"11-14yrs", replace 
label value agecat agecatft 

tab agecat

// Sex
gen sex=. 
replace sex=1 if scr_a5==1|a5==1
replace sex=2 if scr_a5==2|a5==2

label define sexft 1"male" 2"female", replace 
label value sex sexft 

tab sex

// Eligibility 

gen a7new=. 
replace a7new=1 if scr_a72==1&scr_a73==1&scr_a74==1
replace a7new=1 if a72==1&a73==1&a74==1

replace a7new=1 if regexm(scr_a7, "(^1)")
replace a7new=1 if regexm(a7, "(^1)")

replace a7new=2 if scr_a72==1&scr_a73==1&scr_a74==0
replace a7new=2 if a72==1&a73==1&a74==0

replace a7new=3 if regexm(scr_a7, "(5)")
replace a7new=3 if regexm(a7, "(5)")

label define a7newft 1"consent to all+HH visit" 2"consent to all-HH visit" 3"refuse", replace 
label value a7new a7newft 
label var a7new "consent status"

tab a7new

// Date 

gen recruitdate=.
replace recruitdate=date(scr_a2, "YMD") if cohort==1
replace recruitdate=date(a2, "YMD") if cohort==2

format recruitdate %td

browse scr_a2 a2 recruitdate

// Drop duplicates 
sort a1
quietly by a1: gen dup=cond(_N==1,0,_n)

tab dup

drop if a1=="DEM2-PP-0019"&a6=="KH-25"
drop if a1=="DEM1-PP-0073" &i3e=="2025-12-12T14:16:10.213+07:00"

save "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-screening-cohort1&2-cleaned.dta", replace 

// Recruitment by cohort and date 

use "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-screening-cohort1&2-cleaned.dta", clear

tab cohort if recruitdate<date("05/01/2026","DMY") //Nov 20 2025 - Jan 05 2026

tab cohort if recruitdate>=date("05/01/2026","DMY") //Dec 01 2025 - Jan 05 2026

list a1 cohort recruitdate 

/*
 Append lab forms for cohort 1 and 2 
*/

clear 

// Baseline elisa cohort 1 

import excel "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/C1_LabResult_Cohort1_Baseline_ELISA.xlsx", sheet("Sheet 1 - C1_LabResult_Cohort1_") cellrange(A2:BC79) firstrow case(lower)

tostring a1, replace

replace a1="DEM1-PP-0001" if a1manual=="DEM1-PP-0001"
replace a1="DEM1-PP-0002" if a1manual=="DEM1-PP-0002"
replace a1="DEM1-PP-0003" if a1manual=="DEM1-PP-0003"
replace a1="DEM1-PP-0004" if a1manual=="DEM1-PP-0004"
replace a1="DEM1-PP-0005" if a1manual=="DEM1-PP-0005"
replace a1="DEM1-PP-0006" if a1manual=="DEM1-PP-0006"
replace a1="DEM1-PP-0007" if a1manual=="DEM1-PP-0007"
replace a1="DEM1-PP-0008" if a1manual=="DEM1-PP-0008"
replace a1="DEM1-PP-0009" if a1manual=="DEM1-PP-0009"
replace a1="DEM1-PP-0010" if a1manual=="DEM1-PP-0010"
replace a1="DEM1-PP-0011" if a1manual=="DEM1-PP-0011"
replace a1="DEM1-PP-0012" if a1manual=="DEM1-PP-0012"
replace a1="DEM1-PP-0013" if a1manual=="DEM1-PP-0013"
replace a1="DEM1-PP-0014" if a1manual=="DEM1-PP-0014"
replace a1="DEM1-PP-0015" if a1manual=="DEM1-PP-0015"
replace a1="DEM1-PP-0016" if a1manual=="DEM1-PP-0016"
replace a1="DEM1-PP-0017" if a1manual=="DEM1-PP-0017"
replace a1="DEM1-PP-0048" if a1manual=="DEM1-PP-0048"
replace a1="DEM1-PP-0049" if a1manual=="DEM1-PP-0049"
replace a1="DEM1-PP-0047" if a1manual=="DEM1-PP-0047"
replace a1="DEM1-PP-0046" if a1manual=="DEM1-PP-0046"
replace a1="DEM1-PP-0044" if a1manual=="DEM1-PP-0044"
replace a1="DEM1-PP-0043" if a1manual=="DEM1-PP-0043"
replace a1="DEM1-PP-0042" if a1manual=="DEM1-PP-0042"
replace a1="DEM1-PP-0041" if a1manual=="DEM1-PP-0041"
replace a1="DEM1-PP-0040" if a1manual=="DEM1-PP-0040"
replace a1="DEM1-PP-0039" if a1manual=="DEM1-PP-0039"
replace a1="DEM1-PP-0037" if a1manual=="DEM1-PP-0037"
replace a1="DEM1-PP-0036" if a1manual=="DEM1-PP-0036"
replace a1="DEM1-PP-0035" if a1manual=="DEM1-PP-0035"
replace a1="DEM1-PP-0034" if a1manual=="DEM1-PP-0034"
replace a1="DEM1-PP-0033" if a1manual=="DEM1-PP-0033"
replace a1="DEM1-PP-0026" if a1manual=="DEM1-PP-0026"
replace a1="DEM1-PP-0025" if a1manual=="DEM1-PP-0025"
replace a1="DEM1-PP-0024" if a1manual=="DEM1-PP-0024"
replace a1="DEM1-PP-0023" if a1manual=="DEM1-PP-0023"
replace a1="DEM1-PP-0022" if a1manual=="DEM1-PP-0022"
replace a1="DEM1-PP-0021" if a1manual=="DEM1-PP-0021"
replace a1="DEM1-PP-0019" if a1manual=="DEM1-PP-0019"
replace a1="DEM1-PP-0018" if a1manual=="DEM1-PP-0018"

gen a1_5c1new=date(a1_5, "YMD")

format a1_5c1new %td

sort a1
quietly by a1: gen dup2=cond(_N==1,0,_n)

tab dup2

sort dup2

browse a1 dup2

drop if a1=="DEM1-PP-0032"&a1_5c1new<date("25/12/2025","DMY")

save "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-elisaD0-cohort1.dta", replace

// Baseline elisa cohort 2 

clear

import excel "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/C4_LabResult_Cohort2_Baseline_ELISA.xlsx", sheet("Sheet 1 - C4_LabResult_Cohort2_") cellrange(A2:CA21) firstrow case(lower)

replace a1="DEM2-PP-0002" if a1manual=="DEM2-PP-0002"
replace a1="DEM2-PP-0001" if a1manual=="DEM2-PP-0001"
replace a1="DEM2-PP-0003" if a1manual=="DEM2-PP-0003"
replace a1="DEM2-PP-0005" if a1manual=="DEM2-PP-0005"
replace a1="DEM2-PP-0006" if a1manual=="DEM2-PP-0006"
replace a1="DEM2-PP-0007" if a1manual=="DEM2-PP-0007"
replace a1="DEM2-PP-0008" if a1manual=="DEM2-PP-0008"

gen a1_5c2new=date(a1_5, "YMD")
format a1_5c2new %td

sort a1
quietly by a1: gen dup3=cond(_N==1,0,_n)

tab dup3

drop if a1=="DEM2-PP-0006"&a1_5c2new<date("30/12/2025","DMY")
drop if a1=="DEM2-PP-0001"&&a1_5c2new<date("30/12/2025","DMY")
drop if a1=="DEM2-PP-0002"&&a1_5c2new<date("26/12/2025","DMY")
drop if a1=="DEM2-PP-0007"&&a1_5c2new<date("30/12/2025","DMY")

//browse a1 a1manual d0_elic2_b1_5 d0_elic2_b3_5 d0_elic2_c1_5 d0_elic2_c2_5 d0_elic2_c3_5 d0_elic2_c4_5

save "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-elisaD0-cohort2.dta", replace

// Baseline PCR/Qiastat cohort 2

clear

import excel "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/C5_LabResult_Cohort2_Baseline_QIASTAT_PCR.xlsx", sheet("Sheet 1 - C5_LabResult_Cohort2_") cellrange(A2:DI34) firstrow case(lower)

replace a1="DEM2-PP-0008" if a1manual=="DEM2-PP-0008"
replace a1="DEM2-PP-0017" if a1manual=="DEM2-PP-0017"
replace a1="DEM2-PP-0013" if a1manual=="DEM2-PP-0013"
replace a1="DEM2-PP-0007" if a1manual=="DEM2-PP-0007"
replace a1="DEM2-PP-0005" if a1manual=="DEM2-PP-0005"
replace a1="DEM2-PP-0010" if a1manual=="DEM2-PP-0010"

save "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-pcrD0-cohort2.dta", replace

// Baseline Genxpert cohort 2

clear

import excel "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/C6_LabResult_Cohort2_Baseline_ IPC_GenXpert.xlsx", sheet("Sheet 1 - C6_LabResult_Cohort2_") cellrange(A2:AG34) firstrow case(lower)

replace a1="DEM2-PP-0008" if a1manual=="DEM2-PP-0008"
replace a1="DEM2-PP-0006" if a1manual=="DEM2-PP-0006"
replace a1="DEM2-PP-0011" if a1manual=="DEM2-PP-0011"
replace a1="DEM2-PP-0007" if a1manual=="DEM2-PP-0007"
replace a1="DEM2-PP-0005" if a1manual=="DEM2-PP-0005"
replace a1="DEM2-PP-0003" if a1manual=="DEM2-PP-0003"
replace a1="DEM2-PP-0010" if a1manual=="DEM2-PP-0010"

save "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-genxD0-cohort2.dta", replace

// merge elisa, pcr, genX for cohort 2 baseline 

use "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-elisaD0-cohort2.dta", clear 

merge 1:m a1 using "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-pcrD0-cohort2.dta", force 

save "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-elisa&pcrD0-cohort2.dta", replace 

drop _merge

merge 1:m a1 using "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-genxD0-cohort2.dta", force 

drop _merge

save "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-elisa&pcr&genxD0-cohort2.dta", replace 

// Append baseline elisa cohort 1 & 2 and PCR/Qiastat cohort 2

append using "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-elisaD0-cohort1.dta", force

save "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-D0-lab-all-cohort1&2.dta", replace

/* sort a1
quietly by a1: gen dup=cond(_N==1,0,_n)

tab dup */

/*
 Combine screening and lab forms for cohort 1 and 2 - baseline 
*/

clear
 
use "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-D0-lab-all-cohort1&2.dta"

merge 1:m a1 using "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-screening-cohort1&2-cleaned.dta", force 

save "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-D0-screening&lab-cohort1&2.dta", replace 

/*
 Rename and recoding variables 
*/ 

use "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-D0-screening&lab-cohort1&2.dta", clear 

// collect/test date 

gen collectdate=date(a2, "YMD")
gen testdate=date(a1_5, "YMD")

format collectdate testdate %td

// Elisa D0 results cohort 1: IgG anti-Dengue 

label var b5 "D0 serum IgG anti-DEN"

// Elisa D0 results cohort 1: IgG anti-JEV 

label var c5 "D0 serum IgG anti-JEV"

/* 
 Cohort 2
*/

// Elisa D0 results: IgG anti-DEN (serum)

label var d0_elic2_c4_5 "D0 serum IgG anti-DEN"

// Elisa D0 results: IgG anti-JEV (serum)

label var d0_elic2_c2_5 "D0 serum IgG anti-JEV"

// Elisa D0 results: IgM anti-DEN (serum)

label var d0_elic2_c3_5 "D0 serum IgM anti-DEN"

// Elisa D0 results: IgM anti-JEV (serum)

label var d0_elic2_c1_5 "D0 serum IgM anti-JEV"

// Elisa D0 results: IgM anti-JEV (csf)
label var d0_elic2_b1_5 "D0 csf IgM anti-JEV"

// Elisa D0 results: IgM anti-DEN (csf)

label var d0_elic2_b3_5 "D0 csf IgM anti-DEN"

// pcr D0 in csf 
label var d0_pcrc2_b5 "e.coli in csf"
label var d0_pcrc2_b6 "H.Influenzae in csf"
label var d0_pcrc2_b7 "L.monocytogenes in csf"
label var d0_pcrc2_b8 "N.meningitidis in csf"
label var d0_pcrc2_b9 "S.agalactiae in csf"
label var d0_pcrc2_b10 "S.pneumoniae in csf"
label var d0_pcrc2_b11 "M.pneumoniae in csf"
label var d0_pcrc2_b12 "S.pyogenes in csf"
label var d0_pcrc2_b13 "HSV 1 in csf"
label var d0_pcrc2_b14 "HSV 2 in csf"
label var d0_pcrc2_b15 "HPV 6 in csf"
label var d0_pcrc2_b16 "Enterovirus in csf"
label var d0_pcrc2_b17 "Human parechovirus in csf"
label var d0_pcrc2_b18 "VZV in csf"
label var d0_pcrc2_b19 "Cryptococcus gattii/Cryptococcus neoformans in csf"

label var d0_pcrc2_c1_5 "Orientia tsutsugamushi in csf"
label var d0_pcrc2_c1_6 "Rickettsia spp in csf"
label var d0_pcrc2_c1_7 "Leptospira spp in csf"
label var d0_pcrc2_c1_8 "Streptococcus suis in csf"
label var d0_pcrc2_c1_9 "Dengue virus in csf"
label var d0_pcrc2_c1_10 "JEV in csf"

// pcr D0 in serum

label var d0_pcrc2_c2_5 "Orientia tsutsugamushi in serum"
label var d0_pcrc2_c2_6 "Rickettsia spp in serum"
label var d0_pcrc2_c2_7 "Leptospira spp in serum"
label var d0_pcrc2_c2_9 "Dengue virus in serum" 
label var d0_pcrc2_c2_10 "JEV in serum"

label define detectft 1"detected" 2"undetected", replace 
label value d0_pcrc2_b5 d0_pcrc2_b6 d0_pcrc2_b7 d0_pcrc2_b8 d0_pcrc2_b9 d0_pcrc2_b10 ///
	d0_pcrc2_b11 d0_pcrc2_b12 d0_pcrc2_b13 d0_pcrc2_b14 d0_pcrc2_b15 d0_pcrc2_b16 d0_pcrc2_b17 ///
	d0_pcrc2_b18 d0_pcrc2_b19 d0_pcrc2_c1_5 d0_pcrc2_c1_6 d0_pcrc2_c1_7 d0_pcrc2_c1_8 ///
	d0_pcrc2_c1_9 d0_pcrc2_c1_10 d0_pcrc2_c2_5 d0_pcrc2_c2_6 d0_pcrc2_c2_7 d0_pcrc2_c2_9 ///
	d0_pcrc2_c2_10 detectft 
	
// GenXpert baseline 
label var d0_genxc2_b1 "TB in csf"
label define d0_genxc2_b1ft 1"detected high" 2"detected medium" 3"detected low" ///
	4"detected very low" 5"trace detected" 6"not detected" 7"invalid" 8"error" 9"no result", replace 
label value d0_genxc2_b1 d0_genxc2_b1ft

label var d0_genxc2_b2 "Refam resistance"
label define d0_genxc2_b2ft 1"detected" 2"indeterminate" 3"not detected", replace 
label value d0_genxc2_b2 d0_genxc2_b2ft 

// Baseline IgG cohort 1 & 2 

gen iggd0serumjev=. //IgG JEV 
replace iggd0serumjev=1 if b5==1|d0_elic2_c2_5==1
replace iggd0serumjev=2 if b5==2|d0_elic2_c2_5==2
replace iggd0serumjev=3 if b5==3|d0_elic2_c2_5==3

label var iggd0serumjev "D0 serum IgG JEV"

gen iggd0serumden=. //IgG Dengue
replace iggd0serumden=1 if c5==1|d0_elic2_c4_5==1
replace iggd0serumden=2 if c5==2|d0_elic2_c4_5==2
replace iggd0serumden=3 if c5==3|d0_elic2_c4_5==3

label var iggd0serumden "D0 serum IgG DEN"

list cohort a1 iggd0serumden iggd0serumjev //c5 b5 d0_elic2_c4_5 d0_elic2_c2_5

tab iggd0serumjev cohort, chi2 col 

tab iggd0serumden cohort, chi2 col 

// Baseline serum IgM status cohort 2 

gen igmd0serumden=. // Serum IgM dengue 
replace igmd0serumden=1 if d0_elic2_c3_5==1
replace igmd0serumden=2 if d0_elic2_c3_5==2
replace igmd0serumden=3 if d0_elic2_c3_5==3

label var igmd0serumden "D0 serum IgM DEN"

gen igmd0serumjev=. // Serum IgM dengue 
replace igmd0serumjev=1 if d0_elic2_c1_5==1 
replace igmd0serumjev=2 if d0_elic2_c1_5==2
replace igmd0serumjev=3 if d0_elic2_c1_5==3

label var igmd0serumjev "D0 serum IgM JEV"

// Baseline csf IgM status cohort 2 

gen igmd0csfden=.
replace igmd0csfden=1 if d0_elic2_b3_5==1
replace igmd0csfden=2 if d0_elic2_b3_5==2
replace igmd0csfden=3 if d0_elic2_b3_5==3

label var igmd0csfden "D0 csf IgM DEN"

gen igmd0csfjev=.
replace igmd0csfjev=1 if d0_elic2_b1_5==1
replace igmd0csfjev=2 if d0_elic2_b1_5==2
replace igmd0csfjev=3 if d0_elic2_b1_5==3

label var igmd0csfjev "D0 csf IgM JEV"

// Formating IgG and IgM status 
label define poneft 1"pos" 2"neg" 3"equi", replace 
label value b5 c5 d0_elic2_c1_5 d0_elic2_c2_5 d0_elic2_c3_5 d0_elic2_c4_5 d0_elic2_b1_5 ///
	d0_elic2_b3_5 iggd0serumjev iggd0serumden igmd0serumden igmd0serumjev ///
	igmd0csfden igmd0csfjev poneft 

label var a1 "study ID" 

save "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-D0-screening&lab-cohort1&2-cleaned.dta", replace 

// Elisa baseline results by cohort and date 

use "/Users/kennareyseang/Documents/Projects and Proposals/FEF:ANRS/Data collection/Data extracted/Study data/Demele-D0-screening&lab-cohort1&2-cleaned.dta", clear

tab iggd0serumden cohort // IgG dengue status by cohort 

tab iggd0serumdjev cohort // IgG jev status by cohort 

list a1 address iggd0serumden iggd0serumjev if cohort==2 //Baseline serum IgG for cohort 2

list a1 iggd0serumden iggd0serumjev if cohort==1 //Baseline serum IgG for cohort 1

tab iggd0serumden iggd0serumjev if cohort==1

list a1 iggd0serumden if cohort==1&iggd0serumden==1 //Baseline serum IgG anti-DEN (+) for cohort 1

list a1 iggd0serumjev if cohort==1&iggd0serumjev==1 //Baseline serum IgG anti-JEV (+) for cohort 1

//Baseline serum IgM for cohort 2

list a1 address igmd0serumden igmd0serumjev igmd0csfden igmd0csfden if cohort==2

list a1 d0_pcrc2_c2_9 d0_pcrc2_c1_9 if cohort==2 //dengue virus pcr in serum and csf

// TB result at D0 cohort 2 

list d0_genxc2_b1 if cohort==2 

// Recruitment by cohort 
list a1 cohort recruitdate 

tab cohort if recruitdate<date("05/01/2026","DMY") // 20 Nov 225 - 05 Jan 2026











