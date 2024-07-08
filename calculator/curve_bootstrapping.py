import numpy as np
import pandas as pd
from datetime import *
from dateutil.relativedelta import *
from scipy import interpolate

# Class to handle holidays for different calendars
class CountryHoliday:
    def __init__(self):
        pass
    
    # Method to check if a given date is a holiday based on the provided calendar
    def _IsHoliday_(self, date, calendar):
        weekdays = [0, 1, 2, 3, 4]
        if not date.weekday() in weekdays:
            return False
        
        y = date.year
        m = date.month
        d = date.day
        wkd = date.weekday()
       
        if calendar == 'NY':  # Example calendar for New York holidays
            # Check for specific holidays and their observance rules
            if (m == 1 and d == 1) or (m == 12 and d == 31 and wkd == 4) or (m == 1 and d == 2 and wkd == 0):
                return True
            if m == 1 and wkd == 0 and (d > 14 and d < 22):
                return True
            if m == 2 and wkd == 0 and (d > 14 and d < 22):
                return True
            if m == 5 and wkd == 0 and d > 24:
                return True
            if (m == 7 and d == 4) or (m == 7 and d == 3 and wkd == 4) or (m == 7 and d == 5 and wkd == 0):
                return True
            if m == 9 and wkd == 0 and d < 8:
                return True
            if m == 10 and wkd == 0 and (d > 7 and d < 15):
                return True
            if (m == 11 and d == 11) or (m == 11 and d == 10 and wkd == 4) or (m == 11 and d == 12 and wkd == 0):
                return True
            if m == 11 and wkd == 3 and (d > 21 and d < 29):
                return True
            if (m == 12 and d == 25) or (m == 12 and d == 24 and wkd == 4) or (m == 12 and d == 26 and wkd == 0):
                return True
        return False

# Class to handle date calculations for curves, inherits from CountryHoliday
class CurveDate(CountryHoliday):
    def __init__(self):
        pass
    
    # Method to check if a given date is a weekend
    def _IsWeekend_(self, date):
        weekdays = [0, 1, 2, 3, 4]
        return date.weekday() not in weekdays
                       
    # Method to add business days to a given date
    def _AddBusinessDays_(self, date, numdays, busdayconv, calendar):         
        next_date = date
        for i in range(numdays):
            next_date = next_date + relativedelta(days=+1)
            while (self._IsWeekend_(next_date)) or (self._IsHoliday_(next_date, calendar)):
                next_date = next_date + relativedelta(days=+1)
        return next_date
    
    # Method to add business months to a given date
    def _AddBusinessMonths_(self, date, nummonths, busdayconv, calendar):
        next_date = date
        next_date = next_date + relativedelta(months=+nummonths)
        while (self._IsWeekend_(next_date)) or (self._IsHoliday_(next_date, calendar)):
            next_date = next_date + relativedelta(days=+1)
        return next_date
        
    # Method to add business years to a given date
    def _AddBusinessYears_(self, date, numyears, busdayconv, calendar):
        next_date = date
        next_date = next_date + relativedelta(years=+numyears)
        while (self._IsWeekend_(next_date)) or (self._IsHoliday_(next_date, calendar)):
            next_date = next_date + relativedelta(days=+1)
        return next_date
    
    # Method to calculate year fraction between two dates based on day count convention
    def _YFrac_(self, date1, date2, daycountconvention):
        if daycountconvention == 'ACTACT':
            pass
        elif daycountconvention == 'ACT365':
            delta = date2 - date1
            delta_fraction = delta.days / 365.0
            return delta_fraction
        elif daycountconvention == 'ACT360':
            delta = date2 - date1
            delta_fraction = delta.days / 360.0
            return delta_fraction
        elif daycountconvention == 'Thirty360':
            pass
        else:
            delta = date2 - date1
            delta_fraction = delta.days / 360.0
            return delta_fraction

# Class to handle yield curve calculations, inherits from CurveDate and CountryHoliday
class YieldCurve(CurveDate, CountryHoliday):
    def __init__(self):
        self.dfcurve = pd.DataFrame()  # Initialize dataframe for curve data
    
    # Method to read swap curve data from an Excel file
    def __GetSwapCurveData__(self):
        filein = "sofr_data.xlsx"
        self.dfcurve = pd.read_excel(filein, sheet_name='curvedata', index_col='Tenor')
        self.dfcurveparams = pd.read_excel(filein, sheet_name='curveparams')
        self.dfcurveparams.loc[self.dfcurveparams.index[0], 'Date'] = datetime(2024, 7, 2)
        
        busdayconv = self.dfcurveparams['BusDayConv'].iloc[0]
        calendar = self.dfcurveparams['Calendar'].iloc[0]
        curvedate = self.dfcurveparams['Date'].iloc[0]
        
        while (self._IsWeekend_(curvedate) or self._IsHoliday_(curvedate, calendar)):
            curvedate = curvedate + relativedelta(days=+1)
        
        curvesettledays = self.dfcurveparams['SettleDays'].iloc[0]
        curvesettledate = self._AddBusinessDays_(curvedate, curvesettledays, busdayconv, calendar)
        self.dfcurveparams.loc[self.dfcurveparams.index[0], 'SettleDate'] = curvesettledate
    
    # Method to calculate dates for different tenors
    def _DatesForTenors_(self):
        curvesettledate = self.dfcurveparams['SettleDate'].iloc[0]
        busdayconv = self.dfcurveparams['BusDayConv'].iloc[0]
        calendar = self.dfcurveparams['Calendar'].iloc[0]
        
        self.dfcurve.loc[self.dfcurve.index == 'ON', 'Date'] = self._AddBusinessDays_(curvesettledate, 1, busdayconv, calendar)
        self.dfcurve.loc[self.dfcurve.index == '1W', 'Date'] = self._AddBusinessDays_(curvesettledate, 5, busdayconv, calendar)
        
        for i in range(len(self.dfcurve)):
            if self.dfcurve.index[i][-1] == 'D':
                num = int(self.dfcurve.index[i][:-1])
                self.dfcurve.loc[self.dfcurve.index[i], 'Date'] = self._AddBusinessDays_(curvesettledate, num, busdayconv, calendar)
            if self.dfcurve.index[i][-1] == 'M':
                num = int(self.dfcurve.index[i][:-1])
                self.dfcurve.loc[self.dfcurve.index[i], 'Date'] = self._AddBusinessMonths_(curvesettledate, num, busdayconv, calendar)
            elif self.dfcurve.index[i][-1] == 'Y':
                num = int(self.dfcurve.index[i][:-1])
                self.dfcurve.loc[self.dfcurve.index[i], 'Date'] = self._AddBusinessYears_(curvesettledate, num, busdayconv, calendar)
    
    # Method to calculate year fractions for tenors
    def _YearFractionsForTenors_(self):
        curvesettledate = self.dfcurveparams['SettleDate'].iloc[0]
        for i in range(len(self.dfcurve)):
            daycntconv = self.dfcurve['Daycount'].iloc[i]
            if i == 0:
                self.dfcurve.loc[self.dfcurve.index[i], 'YearFraction'] = self._YFrac_(curvesettledate, self.dfcurve['Date'].iloc[i], daycntconv)
            else:
                self.dfcurve.loc[self.dfcurve.index[i], 'YearFraction'] = self._YFrac_(self.dfcurve['Date'].iloc[i - 1], self.dfcurve['Date'].iloc[i], daycntconv)
            self.dfcurve.loc[self.dfcurve.index[i], 'CumYearFraction'] = self._YFrac_(curvesettledate, self.dfcurve['Date'].iloc[i], daycntconv)
    
    # Method to calculate swap year fractions
    def _SwapYearFractions_(self, frequency, term):
        curvesettledate = self.dfcurveparams['SettleDate'].iloc[0]
        busdayconv = self.dfcurveparams['BusDayConv'].iloc[0]
        calendar = self.dfcurveparams['Calendar'].iloc[0]
        
        if frequency == 'S':
            period = 0.5
        elif frequency == 'Q':
            period = 0.25
        else:
            period = 0.5
        
        swap_months = [int(12 * period * i) for i in range(1, 2 * term)]
        swap_dates = [self._AddBusinessMonths_(curvesettledate, nummonths, busdayconv, calendar) for nummonths in swap_months]
        swap_year_fractions = [self._YFrac_(curvesettledate, swap_date, 'ACT/360') for swap_date in swap_dates]
        
        return swap_year_fractions
    
    # Method to calculate zero rates
    def _ZeroRates_(self):
        for i in range(len(self.dfcurve)):
            if self.dfcurve['Type'].iloc[i] == 'Deposit':
                self.dfcurve.loc[self.dfcurve.index[i], 'ZeroRate'] = (1 / self.dfcurve['CumYearFraction'].iloc[i]) * np.log([1.0 + self.dfcurve['Rate'].iloc[i] * self.dfcurve['CumYearFraction'].iloc[i]])
            elif self.dfcurve['Type'].iloc[i] == 'EuroDollarFuture':
                rate_continuous = 4 * np.log([1.0 + self.dfcurve['Rate'].iloc[i] * self.dfcurve['YearFraction'].iloc[i]])
                self.dfcurve.loc[self.dfcurve.index[i], 'ZeroRate'] = (rate_continuous * self.dfcurve['YearFraction'].iloc[i] + self.dfcurve['ZeroRate'].iloc[i - 1] * self.dfcurve['CumYearFraction'].iloc[i - 1]) / self.dfcurve['CumYearFraction'].iloc[i]
            else:
                sumproduct = 0.0
                
                if self.dfcurve['Type'].iloc[i] == 'Swap':
                    frequency = self.dfcurve['Frequency'].iloc[i]
                    term = int(self.dfcurve.index[i][:-1])
                    swap_year_fractions = self._SwapYearFractions_(frequency, term)
                    
                    # Set up interpolation object
                    x = pd.Series(self.dfcurve['CumYearFraction'][:i])
                    y = pd.Series(self.dfcurve['ZeroRate'][:i])
                    zero_tck = interpolate.splrep(x, y)
                    
                for swap_yf in swap_year_fractions:
                    zero_rate = interpolate.splev(swap_yf, zero_tck)
                    sumproduct += (self.dfcurve['Rate'].iloc[i] / 2.0) * np.exp(-zero_rate * swap_yf)
                
                self.dfcurve.loc[self.dfcurve.index[i], 'ZeroRate'] = (-1 * np.log((1.0 - sumproduct) / (1.0 + self.dfcurve['Rate'].iloc[i] / 2.0))) / self.dfcurve['CumYearFraction'].iloc[i]
        
        # Save the calculated yield curve to an Excel file
        fileout = "yieldcurve.xlsx"
        self.dfcurve.to_excel(fileout, sheet_name='yieldcurve', index=True)
    
    # Method to calculate discount factors
    def _DiscountFactors_(self):
        for i in range(len(self.dfcurve)):
            rate_i = self.dfcurve['Rate'].iloc[i]
            yearfraction_i = self.dfcurve['YearFraction'].iloc[i]
            cumyearfraction_i = self.dfcurve['CumYearFraction'].iloc[i]
            self.dfcurve.loc[self.dfcurve.index[i], 'DiscountFactor'] = np.exp(-rate_i * cumyearfraction_i)
    
    # Method to calculate forward rates
    def _ForwardRates_(self):
        for i in range(len(self.dfcurve)):
            if i == 0:
                self.dfcurve.loc[self.dfcurve.index[i], 'ForwardRate'] = 0.0
            else:
                discountfactor_i = self.dfcurve['DiscountFactor'].iloc[i]
                discountfactor_iminusone = self.dfcurve['DiscountFactor'].iloc[i - 1]
                yearfraction_i = self.dfcurve['YearFraction'].iloc[i]
                self.dfcurve.loc[self.dfcurve.index[i], 'ForwardRate'] = (discountfactor_iminusone / discountfactor_i - 1) / yearfraction_i
    
    # Method to bootstrap the yield curve
    def BootstrapYieldCurve(self):
        self.__GetSwapCurveData__()
        self._DatesForTenors_()
        self._YearFractionsForTenors_()
        self._ZeroRates_()
        self._DiscountFactors_()
        self._ForwardRates_()