cls
clear

capture log close

set linesize 90
set scheme s2mono

graph drop _all
set graphics on

// alter the path to where the excel file is
local path = "C:\Users\Legion\Desktop\market-efficiency-analysis"

if "`path'" == "" {
	di "Alter the path variable to the directory where the excel is"
	exit
}

cd "`path'"

// if save = 1 saves all matrix´s
local save = 1

// foreach sheet name
foreach index in "CAC" "DAX" "DJ" "EURSTOXX" "FTSE" "HSI" "NASDAQ" "NIKKEI" "SMI" "SP500" "SSEC" "PSI20" {
	
	di "Starting `index'"
	
	// Data importing
	import excel "stock_market_index_data.xlsx", sheet("`index'") firstrow allstring
	
	// converting the prices to numeric variable
	destring `index', generate(prices)
	drop `index'
	drop if prices == .

	// formating the time
	gen dates = date(DATAX, "MDY")
	format %td dates
	drop DATAX
	
	// for different time intervals
	// daily
	gen daily = 1
	
	// weekly 
	destring days_of_week, generate(week)
	gen weekly = .
	replace weekly = 1 if week == 4
	drop week
	
	// monthly
	destring end_month, generate(monthly)
	drop end_month
	
	// "daily" "weekly" "monthly"
	foreach interval in "daily" "weekly" "monthly" {
		preserve

		keep if `interval' == 1
		
		// returns
		gen ln_index = ln(prices)
		gen returns = ln_index - ln_index[_n - 1]
		drop if returns == .
	
		// generating time variable
		gen time = _n
		tsset time
		
		// price and returns over time
		line prices dates, ytitle("Prices") xtitle("Time") title("`index' `interval' prices over time") name("`index'_`interval'_prices_time")
		line returns dates, ytitle("Returns") xtitle("Time") title("`index' `interval' returns over time") name("`index'_`interval'_returns_time")
		
		// returns histogram
		histogram returns, normal ytitle("Density") xtitle("Returns") title("`index' `interval' histogram returns") name("`index'_`interval'_hist_returns")
		
		// Jarque-Bera test
		quietly: tabstat returns, stat(n mean sd min max skewness kurtosis) save
		matrix `index'_`interval'_statistical = r(StatTotal)'
		quietly: sktest returns
		local jb = (`index'_`interval'_statistical[1,1] / 6) * (`index'_`interval'_statistical[1,6] ^ 2 + 0.25 * (`index'_`interval'_statistical[1,7] - 3)^2)
		local jb_pvalue = r(p_chi2)
		matrix jb_test = (`jb',`jb_pvalue')
		matrix `index'_`interval'_statistical = `index'_`interval'_statistical, jb_test
		matrix colnames `index'_`interval'_statistical = "N" "Mean" "Standard-Deviation" "Min" "Max" "Skewness" "Kurtosis" "Jarque-Bera" "p-value"
		matrix rownames `index'_`interval'_statistical = "`index' `interval'"
		
		// autocorrelation tests
		ac returns, title("Autocorrelation of `index' `interval' returns") name("`index'_ac_`interval'_returns")

		foreach n of numlist 1/10 {
			quietly: wntestq returns, lag(`n')
			
			if `n' == 1 {
				matrix `index'_`interval'_autocorrelation = (r(stat) \ r(p))
				matrix rownames `index'_`interval'_autocorrelation = "`index' `interval' Q statistic" "`index' `interval' p-value"
				matrix colnames `index'_`interval'_autocorrelation = "lag `n'"
			}
			else {
				matrix helper = (r(stat) \ r(p))
				matrix colnames helper = "lag `n'"
				matrix `index'_`interval'_autocorrelation = `index'_`interval'_autocorrelation, helper
			}
		}
		 
		// unit root test - ADF
		quietly: varsoc returns time
		local optimal_lag = r(mlag)
		quietly: dfuller returns, lags(`optimal_lag')

		matrix `index'_`interval'_adf_test = (r(N), `optimal_lag', r(Zt), r(p))
		matrix colnames `index'_`interval'_adf_test = "N" "Lag" "ADF Statistic" "p-value"
		matrix rownames `index'_`interval'_adf_test = "`index' `interval'"
		
		// variance ratio test - Lomackinlay
		quietly: lomackinlay ln_index
		
		matrix `index'_`interval'_ratio_test = (r(v_2) \ r(R_2) \ r(p_2)), (r(v_4) \ r(R_4) \ r(p_4)), (r(v_8) \ r(R_8) \ r(p_8)), (r(v_16) \ r(R_16) \ r(p_16))
		matrix colnames `index'_`interval'_ratio_test = "lag 2" "lag 4" "lag 8" "lag 16"
		matrix rownames `index'_`interval'_ratio_test = "`index' `interval' VR" "`index' `interval' R_s" "`index' `interval' p-value"
		
		// runs test 
		quietly: runtest returns, m
		
		matrix `index'_`interval'_runs_test = (r(N), r(N_below), r(N_above), r(z), r(p))
		matrix colnames `index'_`interval'_runs_test = "N" "N <= mean" "N > mean" "Z" "p-value"
		matrix rownames `index'_`interval'_runs_test = "`index' `interval'"
		
		restore
	}
	
	
	matrix final_`index'_statistical = `index'_daily_statistical \ `index'_weekly_statistical \ `index'_monthly_statistical
	matrix final_`index'_autocorrelation = `index'_daily_autocorrelation \ `index'_weekly_autocorrelation \ `index'_monthly_autocorrelation
	matrix final_`index'_adf_test = `index'_daily_adf_test \ `index'_weekly_adf_test \ `index'_monthly_adf_test
	matrix final_`index'_ratio_test = `index'_daily_ratio_test \ `index'_weekly_ratio_test \ `index'_monthly_ratio_test
	matrix final_`index'_runs_test = `index'_daily_runs_test \ `index'_weekly_runs_test \ `index'_monthly_runs_test
	
	drop YEAR-weekly
}


foreach index in "CAC" "DAX" "DJ" "EURSTOXX" "FTSE" "HSI" "NASDAQ" "NIKKEI" "SMI" "SP500" "SSEC" "PSI20" {
	if "`index'" == "CAC" {
		matrix final_statistical =  final_`index'_statistical
		matrix final_autocorrelation =  final_`index'_autocorrelation
		matrix final_adf_test =  final_`index'_adf_test
		matrix final_ratio_test =  final_`index'_ratio_test
		matrix final_runs_test =  final_`index'_runs_test
	}
	else {
		matrix final_statistical = final_statistical \ final_`index'_statistical
		matrix final_autocorrelation =  final_autocorrelation \ final_`index'_autocorrelation
		matrix final_adf_test =  final_adf_test \ final_`index'_adf_test
		matrix final_ratio_test =  final_ratio_test \ final_`index'_ratio_test
		matrix final_runs_test =  final_runs_test \ final_`index'_runs_test
	}
}

if `save' == 1 {
	
	capture mkdir excels

	putexcel set "excels/statistical_summary.xlsx", replace
	putexcel A1 = matrix(final_statistical), names
	putexcel set "excels/autocorrelation.xlsx", replace
	putexcel A1 = matrix(final_autocorrelation), names
	putexcel set "excels/unit_root_adf.xlsx", replace
	putexcel A1 = matrix(final_adf_test), names
	putexcel set "excels/ratio_test_lomackinlay.xlsx", replace
	putexcel A1 = matrix(final_ratio_test), names
	putexcel set "excels/runs_test.xlsx", replace
	putexcel A1 = matrix(final_runs_test), names
	
	capture mkdir graphs
	
	foreach index in "CAC" "DAX" "DJ" "EURSTOXX" "FTSE" "HSI" "NASDAQ" "NIKKEI" "SMI" "SP500" "SSEC" "PSI20" {
		
		capture mkdir graphs/`index'
		
		foreach interval in "daily" "weekly" "monthly" {
			graph export "graphs/`index'/`index'_`interval'_prices_time.png", replace name("`index'_`interval'_prices_time")
			graph export "graphs/`index'/`index'_`interval'_returns_time.png", replace name("`index'_`interval'_returns_time")
			graph export "graphs/`index'/`index'_`interval'_hist_returns.png", replace name("`index'_`interval'_hist_returns")
			graph export "graphs/`index'/`index'_ac_`interval'_returns.png", replace name("`index'_ac_`interval'_returns")
		}
	}
}

matlist final_statistical
matlist final_autocorrelation
matlist final_adf_test
matlist final_ratio_test
matlist final_runs_test
