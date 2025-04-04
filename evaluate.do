global START_MONTH = "05"
global PATH = "Jobs.xlsx"

/*******************************************************************************
	2024 Macro
********************************************************************************/
import excel "${PATH}", sheet("Macro_2024_2025") firstrow clear
replace PostDate = "9/17/2024" if PostDate == "Reposted 3/27/2025 (originally 9/17/2024)"
replace PostDate = "8/4/2024" if PostDate == "Reposted 2/11/2025 (originally 8/4/2024)"
replace PostDate = "10/24/2024" if PostDate == "Reposted 2/11/2035 (originally 10/24/2024)"
drop if strpos(PostDate, "2020")
drop if strpos(PostDate, "2023")
drop if missing(PostDate)
replace PostDate = trim(PostDate)
gen date = date(PostDate, "MDY")
format %td date
assert !missing(date)
order date, first
drop if date < date("${START_MONTH}/01/2024", "MDY")
gen rel_date = date - date("08/09/2024", "MDY")
gcollapse (count) macro24=date, by(rel_date)
tempfile t
save "`t'", replace

/*******************************************************************************
	2024 Micro
********************************************************************************/
import excel "${PATH}", sheet("Micro_2024_2025") firstrow clear
rename Date PostDate
replace PostDate = "10/24/2024" if PostDate == "Reposted 2/11/2025 (originally 10/24/2024)"
replace PostDate = "10/22/2024" if PostDate == "Reposted 3/25/2025 (originally 10/22/2024)"
replace PostDate = "10/4/2024" if PostDate == "Reposted 2/11/2025 (originally 10/4/2024)"
replace PostDate = "9/30/2024" if PostDate == "Reposted 2/24/25 (originally 9/30/2024)"
replace PostDate = "9/17/2024" if PostDate == "Reposted 3/27/2025 (originally 9/17/2024)"
replace PostDate = "8/29/2024" if PostDate == "Reposted 2/26/2025 (originally 8/29/2024)"
replace PostDate = "2/14/2025" if PostDate == "Feb-14"
replace PostDate = "2/13/2025" if PostDate == "Feb-13"
drop if strpos(PostDate, "2020")
drop if strpos(PostDate, "2023")
drop if missing(PostDate)
replace PostDate = trim(PostDate)
gen date = date(PostDate, "MDY")
format %td date
assert !missing(date)
order date, first
drop if date < date("${START_MONTH}/01/2024", "MDY")
gen rel_date = date - date("08/09/2024", "MDY")
gcollapse (count) micro24=date, by(rel_date)
merge 1:1 rel_date using "`t'"
drop _merge
save "`t'", replace

/*******************************************************************************
	2023 Macro
********************************************************************************/
import excel "${PATH}", sheet("Macro_2023_2024") firstrow clear
drop if missing(Posted)
rename Posted PostDate
replace PostDate = trim(PostDate)
gen date = date(PostDate, "MDY")
format %td date
assert !missing(date)
order date, first
drop if date < date("${START_MONTH}/01/2023", "MDY")
gen rel_date = date - date("08/05/2023", "MDY")
gcollapse (count) macro23=date, by(rel_date)
merge 1:1 rel_date using "`t'"
drop _merge
save "`t'", replace

/*******************************************************************************
	2023 Micro
********************************************************************************/
import excel "${PATH}", sheet("Micro_2023_2024") firstrow clear
rename Date PostDate
replace PostDate = trim(PostDate)
drop if _n >= 274
replace PostDate = PostDate[_n-1] if missing(PostDate)
drop if missing(PostDate)
gen date = date(PostDate, "MDY")
format %td date
order date, first
drop if date < date("${START_MONTH}/01/2023", "MDY")
gen rel_date = date - date("08/05/2023", "MDY")
gcollapse (count) micro23=date, by(rel_date)
merge 1:1 rel_date using "`t'"
drop _merge
save "`t'", replace

/*******************************************************************************
	2022 Macro
********************************************************************************/
import excel "${PATH}", sheet("Macro_2022_2023") firstrow clear
destring r, replace
replace PostDate = subinstr(PostDate, "/22", "/2022", .)
replace PostDate = subinstr(PostDate, "/2022/", "/22/", .)
replace PostDate = trim(PostDate)
replace PostDate = PostDate[_n-1] if missing(PostDate)
forvalues i = 1/12 {
	replace PostDate = "5/`i'/2022" if PostDate == "`i'/5/2022" & r < 15
	replace PostDate = "6/`i'/2022" if PostDate == "`i'/6/2022" & r > 15 & r < 75
	replace PostDate = "7/`i'/2022" if PostDate == "`i'/7/2022" & r > 60 & r < 90
}
gen date = date(PostDate, "MDY")
format %td date
order date, first
drop if date < date("${START_MONTH}/01/2022", "MDY")
gen rel_date = date - date("08/05/2022", "MDY")
gcollapse (count) macro22=date, by(rel_date)
merge 1:1 rel_date using "`t'"
drop _merge
save "`t'", replace

/*******************************************************************************
	2022 Micro
********************************************************************************/
import excel "${PATH}", sheet("Micro_2022_2023") firstrow clear
replace r = "37" if r == "B"
destring r, replace
replace PostDate = subinstr(PostDate, "19/23", "19/2022", .)
replace PostDate = subinstr(PostDate, "/22", "/2022", .)
replace PostDate = subinstr(PostDate, "/2022/", "/22/", .)
replace PostDate = trim(PostDate)
replace PostDate = PostDate[_n-1] if missing(PostDate)
forvalues i = 1/12 {
	replace PostDate = "5/`i'/2022" if PostDate == "`i'/5/2022" & r < 15
	replace PostDate = "6/`i'/2022" if PostDate == "`i'/6/2022" & r > 15 & r < 75
	replace PostDate = "7/`i'/2022" if PostDate == "`i'/7/2022" & r > 60 & r < 90
}
gen date = date(PostDate, "MDY")
format %td date
order date, first
drop if date < date("${START_MONTH}/01/2022", "MDY")
gen rel_date = date - date("08/05/2022", "MDY")
gcollapse (count) micro22=date, by(rel_date)
merge 1:1 rel_date using "`t'"
drop _merge
save "`t'", replace

/*******************************************************************************
	2021 Macro
********************************************************************************/
import excel "${PATH}", sheet("Macro_2021_2022") firstrow clear
gsort -r
replace PostDate = trim(PostDate)
forvalues i = 1/12 {
	replace PostDate = "5/`i'/2021" if PostDate == "`i'/5/2021" & r < 15
	replace PostDate = "6/`i'/2021" if PostDate == "`i'/6/2021" & r > 25 & r < 75
	replace PostDate = "7/`i'/2021" if PostDate == "`i'/7/2021" & r > 90 & r < 145
	replace PostDate = "8/`i'/2021" if PostDate == "`i'/8/2021" & r > 180 & r < 230
	replace PostDate = "9/`i'/2021" if PostDate == "`i'/9/2021" & r > 240 & r < 260
	replace PostDate = "10/`i'/2021" if PostDate == "`i'/10/2021" & r > 290 & r < 320
	replace PostDate = "11/`i'/2021" if PostDate == "`i'/11/2021" & r > 330 & r < 365
}
gen date = date(PostDate, "MDY")
format %td date
drop if date < date("${START_MONTH}/01/2021", "MDY")
gen rel_date = date - date("07/30/2021", "MDY")
gcollapse (count) macro21=date, by(rel_date)
merge 1:1 rel_date using "`t'"
drop _merge
save "`t'", replace

/*******************************************************************************
	2021 Macro
********************************************************************************/
import excel "${PATH}", sheet("Micro_2021_2022") firstrow clear
replace PostDate = trim(PostDate)
replace PostDate = subinstr(PostDate, "/20222", "/2022", .)
replace PostDate = PostDate[_n-1] if missing(PostDate)
gen date = date(PostDate, "MDY")
format %td date
order date, first
drop if date < date("${START_MONTH}/01/2021", "MDY")
gen rel_date = date - date("07/30/2021", "MDY")
gcollapse (count) micro21=date, by(rel_date)
merge 1:1 rel_date using "`t'"
drop _merge
save "`t'", replace

/*******************************************************************************
	2020 Macro
********************************************************************************/
import excel "${PATH}", sheet("Macro_2020_2021") firstrow clear
drop r
gen r = _n
order r, first
replace PostDate = trim(PostDate)
forvalues i = 1/12 {
	replace PostDate = "5/`i'/2020" if PostDate == "`i'/5/2020" & r < 15
	replace PostDate = "6/`i'/2020" if PostDate == "`i'/6/2020" & r > 5 & r < 30
	replace PostDate = "7/`i'/2020" if PostDate == "`i'/7/2020" & r > 30 & r < 50 
	replace PostDate = "8/`i'/2020" if PostDate == "`i'/8/2020" & r > 65 & r < 80
	replace PostDate = "9/`i'/2020" if PostDate == "`i'/9/2020" & r > 80 & r < 100
	replace PostDate = "11/`i'/2020" if PostDate == "`i'/11/2020" & r > 120 & r < 137
	replace PostDate = "12/`i'/2020" if PostDate == "`i'/12/2020" & r > 150 & r < 190
}
drop if r > 180
gen date = date(PostDate, "MDY")
format %td date
drop if date < date("${START_MONTH}/01/2020", "MDY")
gen rel_date = date - date("08/07/2020", "MDY")
gcollapse (count) macro20=date, by(rel_date)
merge 1:1 rel_date using "`t'"
drop _merge
save "`t'", replace

/*******************************************************************************
	2020 Micro
********************************************************************************/
import excel "${PATH}", sheet("Micro_2020_2021") firstrow clear
replace PostDate = trim(PostDate)
gen date = date(PostDate, "MDY")
format %td date
order date, first
gen rel_date = date - date("08/07/2020", "MDY")
order rel_date, first
drop if date < date("${START_MONTH}/01/2020", "MDY")
sort rel_date
gcollapse (count) micro20=date, by(rel_date)
merge 1:1 rel_date using "`t'"
drop _merge
save "`t'", replace

/*******************************************************************************
	2019 Macro
********************************************************************************/
import excel "${PATH}", sheet("Macro_2019_2020") firstrow clear
replace Date = trim(Date)
forvalues i = 1/12 {
	replace Date = "11/`i'/2019" if Date == "`i'/11/2019" & _n < 70
	replace Date = "10/`i'/2019" if Date == "`i'/10/2019" & _n < 70
	replace Date = "9/`i'/2019" if Date == "`i'/9/2019" & _n > 80
	replace Date = "8/`i'/2019" if Date == "`i'/8/2019" & _n > 110
	replace Date = "7/`i'/2019" if Date == "`i'/7/2019" & _n > 150
	replace Date = "6/`i'/2019" if Date == "`i'/6/2019" & _n > 150
	replace Date = "5/`i'/2019" if Date == "`i'/5/2019" & _n > 150
}
gen date = date(Date, "MDY")
replace date = date(Date, "DMY") if missing(date)
format %td date
order date, first
drop if date < date("${START_MONTH}/01/2019", "MDY")
gen rel_date = date - date("08/09/2019", "MDY")
gcollapse (count) macro19=date, by(rel_date)
merge 1:1 rel_date using "`t'"
drop _merge
save "`t'", replace

/*******************************************************************************
	2019 Micro
********************************************************************************/
import excel "${PATH}", sheet("Micro_2019_2020") firstrow clear
rename A Date
replace Date = trim(Date)
replace Date = "3/9/2020" if Date == "3/9/20"
gen date = date(Date, "MDY")
format %td date
order date, first
drop if date < date("${START_MONTH}/01/2019", "MDY")
gen rel_date = date - date("08/09/2019", "MDY")
gcollapse (count) micro19=date, by(rel_date)
merge 1:1 rel_date using "`t'"
drop _merge
save "`t'", replace

/*******************************************************************************
	2018 Macro
********************************************************************************/
import excel "${PATH}", sheet("Macro_2018_2019") firstrow clear
replace Date = trim(Date)
forvalues i = 1/12 {
	replace Date = "11/`i'/2018" if Date == "`i'/11/2018" & _n < 60
	replace Date = "10/`i'/2018" if Date == "`i'/10/2018" & _n < 60
	replace Date = "9/`i'/2018" if Date == "`i'/9/2018" & _n < 80
	replace Date = "8/`i'/2018" if Date == "`i'/8/2018" & _n < 150
	replace Date = "7/`i'/2018" if Date == "`i'/7/2018" & _n > 150
	replace Date = "6/`i'/2018" if Date == "`i'/6/2018" & _n > 200
	replace Date = "5/`i'/2018" if Date == "`i'/5/2018" & _n > 240
}
gen date = date(Date, "MDY")
format %td date
order date, first
drop if date < date("${START_MONTH}/01/2018", "MDY")
gen rel_date = date - date("08/10/2018", "MDY")
gcollapse (count) macro18=date, by(rel_date)
merge 1:1 rel_date using "`t'"
drop _merge
save "`t'", replace

/*******************************************************************************
	2018 Micro
********************************************************************************/
import excel "${PATH}", sheet("Micro_2018_2019") firstrow clear
replace Date = trim(Date)
gen date = date(Date, "MDY")
format %td date
order date, first
drop if date < date("${START_MONTH}/01/2018", "MDY")
gen rel_date = date - date("08/10/2018", "MDY")
order rel_date, first
gcollapse (count) micro18=date, by(rel_date)
merge 1:1 rel_date using "`t'"
drop _merge
save "`t'", replace

/*******************************************************************************
	Final processing & graphs
********************************************************************************/
mvencode micro* macro*, mv(.=0) override
sort rel_date
foreach vari of varlist micro* macro* {
	gen cum_`vari' = sum(`vari')
	local lbl = subinstr("`vari'", "macro", "", .)
	local lbl = subinstr("`lbl'", "micro", "", .)
	label var cum_`vari' "AOM 20`lbl'"
}
replace cum_macro23 = . if rel_date > today() - date("08/05/2023", "MDY")
replace cum_micro23 = . if rel_date > today() - date("08/05/2023", "MDY")
drop if rel_date > 200
label var rel_date "Date relative to AOM"

cap drop scatter_today
gen scatter_today = cum_macro24
label var scatter_today "`=trim(c(current_date))'"
twoway ///
	(line cum_macro18 rel_date) ///
	(line cum_macro19 rel_date) ///
	(line cum_macro20 rel_date) ///
	(line cum_macro21 rel_date) ///
	(line cum_macro22 rel_date) ///
	(line cum_macro23 rel_date) ///
	(line cum_macro24 rel_date) ///
	/// (scatter scatter_today rel_date if rel_date == today() - date("07/25/2025", "MDY"), ///
	/// 	msize(small) msymbol(X) mcolor(black)) ///
	, xline(0) scheme(white) ///
	legend(position(10)) title("Job announcements in Macro") ///
	xlabel(-30 "-1M" -60 "-2M" -90 "-3M" 0 "AOM" 30 "+1M" 60 "+2M" 90 "+3M" 150 "New Year")
graph export "AOM_Jobs_Macro.png", replace
	
cap drop scatter_today
gen scatter_today = cum_micro24
label var scatter_today "`=trim(c(current_date))'"
twoway ///
	(line cum_micro18 rel_date) ///
	(line cum_micro19 rel_date) ///
	(line cum_micro20 rel_date) ///
	(line cum_micro21 rel_date) ///
	(line cum_micro22 rel_date) ///
	(line cum_micro23 rel_date) ///
	(line cum_micro24 rel_date) ///
	/// (scatter scatter_today rel_date if rel_date == today() - date("07/25/2025", "MDY"), ///
	///	msize(small) msymbol(X) mcolor(black)) ///
	, xline(0) scheme(white) ///
	legend(position(10)) title("Job announcements in Micro") ///
	xlabel(-30 "-1M" -60 "-2M" -90 "-3M" 0 "AOM" 30 "+1M" 60 "+2M" 90 "+3M" 150 "New Year")
graph export "AOM_Jobs_Micro.png", replace
