// Breaking Through to .NET - Example 1 - Basic COM interop
// March 2014
// Paul Voelker

#include "stdafx.h"

#include "decimal.h" // Wrapper for DECIMAL structure (from http://www.codeproject.com/Articles/4161/DECIMAL-Wrapper-Class)
#include <ATLComTime.h>
#include "COMHelpers.h"

// This creates the file 'currencyconverterbusinesslogic.tlh'
#import "..\TypeLibs\CurrencyConverter.tlb" no_namespace

int _tmain(int argc, _TCHAR* argv[])
{
	CoInitialize(NULL);

	try
	{
		// Instantiate .NET class as a COM object
		ICurrencyConverterPtr currencyConverter(__uuidof(CurrencyConverter));

		// Test 1 - Convert US Dollar to Euro
		// --------------------------------------------------------------------
		_tprintf(_T("+------------------------------------+\n"));
		_tprintf(_T("| Test 1 - Convert US Dollar to Euro |\n"));
		_tprintf(_T("+------------------------------------+\n"));

		COleDateTime exchangeRateDateTime = currencyConverter->GetExchangeRateDateTime(); 
		CString exchangeRateDateTimeStr = exchangeRateDateTime.Format(_T("%A, %B %d, %Y"));

		CString currenyNameUSDollar =
			currencyConverter->GetCurrencyName(CurrencyType_USDollar);
		CString currenyNameEuro =
			currencyConverter->GetCurrencyName(CurrencyType_Euro);

		com::Decimal haveCurrencyValue = 12.34;
		com::Decimal needCurrencyValue =
			currencyConverter->ConvertCurrency(CurrencyType_USDollar,
			haveCurrencyValue, CurrencyType_Euro);
		CString needCurrencyValueStr = needCurrencyValue.ToString();

		_tprintf(_T("On %s,\n      %s in '%s'\n   converts to\n      %s in '%s'\n\n"),
			(LPCTSTR)exchangeRateDateTimeStr,
			(LPCTSTR)(haveCurrencyValue.ToString()), (LPCTSTR)currenyNameUSDollar,
			(LPCTSTR)(needCurrencyValue.ToString()), (LPCTSTR)currenyNameEuro);

		// Test 2 - Throw error inside of .NET
		// --------------------------------------------------------------------
		_tprintf(_T("+-------------------------------------+\n"));
		_tprintf(_T("| Test 2 - Throw Error Inside of .NET |\n"));
		_tprintf(_T("+-------------------------------------+\n"));		

		 // Pass in invalid 'CurrencyType' which will cause a .NET exception to be thrown
		needCurrencyValue =
			currencyConverter->ConvertCurrency(CurrencyType_USDollar,
			haveCurrencyValue, (CurrencyType)9999999999);

		// NOTE: 'Debug.WriteLine' from .NET code should appear in debug output window
	}
	catch(_com_error& ex)
	{
		_tprintf(_T("COM Error: %s - %s\n"), ex.ErrorMessage(),
			GetCOMError(ex.ErrorInfo()));
	}

	CoUninitialize();

	_tprintf(_T("\n"));

	::system("PAUSE");

	return 0;
}

