// Breaking Through to .NET - Example 2 - COM interop collections
// March 2014
// Paul Voelker

#include "stdafx.h"
#include <atlstr.h>
#include "decimal.h" // Wrapper for DECIMAL structure (from http://www.codeproject.com/Articles/4161/DECIMAL-Wrapper-Class)

#import "C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727\mscorlib.tlb" rename("ReportEvent", "DotNetReportEvent")
#import "..\TypeLibs\Denominations.tlb" no_namespace

// ----------------------------------------------------------------------------
// Timing tests function
// ----------------------------------------------------------------------------
#pragma optimize( "", off ) //  Make sure optimizations are OFF
void TimingTests(IDenominationManagerPtr denomMgr)
{
	long weaklyTypedTicks;
	long stronglyTypedTicks;
	long nativeDotNetTicks;

	// 500000 element collection with strings of 10 characters each
	const int collectionSize = 500000;
	const int nameSize = 10;
	denomMgr->Test_GenerateRandomData(collectionSize, nameSize);

	_tprintf(_T("Test data generated:\n   %d elements with strings of %d characters\n\n"),
		collectionSize, nameSize);

	_tprintf(_T("Testing in progress...\n\n"));

	long startTickCount = ::GetTickCount();
	{
		mscorlib::_ArrayListPtr weakDenomList = denomMgr->GetDenominationsWeakColl();
		
		mscorlib::IListPtr listPtr;
		weakDenomList->QueryInterface(__uuidof(mscorlib::IList), (void **)&listPtr);

		mscorlib::ICollectionPtr collectionPtr;
		weakDenomList->QueryInterface(__uuidof(mscorlib::ICollection), (void **)&collectionPtr);

		const long count = collectionPtr->GetCount();

		for(int i = 0; i < count; i++)
		{
			_variant_t varDenom = listPtr->GetItem(i);
			IDenominationPtr denom = varDenom;
		}
	}
	weaklyTypedTicks = ::GetTickCount() - startTickCount;

	startTickCount = ::GetTickCount();
	{
		IDenominationListCOMPtr strongDenomList = denomMgr->GetDenominationsStrongColl(); 

		const long count = strongDenomList->GetCount();

		for(int i = 0; i < count; i++)
		{
			IDenominationPtr denom = strongDenomList->GetItem(i);
		}
	}
	stronglyTypedTicks = ::GetTickCount() - startTickCount;

	denomMgr->raw_Test_TimeEnumerateDenominations(&nativeDotNetTicks);

	_tprintf(_T("Weakly Typed Collection Time: %d ms\n"), weaklyTypedTicks);
	_tprintf(_T("Strongly Typed Collection Time: %d ms\n"), stronglyTypedTicks);
	_tprintf(_T("Native .NET Time: %d ms\n"), nativeDotNetTicks);
}
#pragma optimize( "", on ) // Turn optimizations back on

// ----------------------------------------------------------------------------
// Main
// ----------------------------------------------------------------------------
int _tmain(int argc, _TCHAR* argv[])
{
	CoInitialize(NULL); // Must be called to use COM

	try
	{
		IDenominationManagerPtr denomMgr(__uuidof(DenominationManager));

		//denomMgr->raw_Test_GenerateRandomData(100, 1);

		// Weakly typed collection
		// --------------------------------------------------------------------------
		{
			_tprintf(_T("+-------------------------+\n"));
			_tprintf(_T("| Weakly Typed Collection |\n"));
			_tprintf(_T("+-------------------------+\n\n"));

			mscorlib::_ArrayListPtr weakDenomList = denomMgr->GetDenominationsWeakColl();
		
			mscorlib::IListPtr listPtr;
			weakDenomList->QueryInterface(__uuidof(mscorlib::IList), (void **)&listPtr);

			mscorlib::ICollectionPtr collectionPtr;
			weakDenomList->QueryInterface(__uuidof(mscorlib::ICollection), (void **)&collectionPtr);

			const long count = collectionPtr->GetCount();
			_tprintf(_T("Number of Denominations: %d\n\n"), count);

			for(int i = 0; i < count; i++)
			{
				_variant_t varDenom = listPtr->GetItem(i); // Weakly typed collection returns variant type
				IDenominationPtr denom = varDenom; // Query interface (via COM smart pointer) for specific data type
				com::Decimal value = denom->GetValue();
				_tprintf(_T("Denom #%d - Name = %s, Value = %s, Is Coin = %s\n"), i + 1,
					denom->GetName().GetBSTR(),
					value.ToString(),
					denom->GetIsCoin() ? _T("True") : _T("False"));
			}
		}

		_tprintf(_T("\n"));

		// Strongly typed collection
		// --------------------------------------------------------------------------
		{
			_tprintf(_T("+---------------------------+\n"));
			_tprintf(_T("| Strongly Typed Collection |\n"));
			_tprintf(_T("+---------------------------+\n\n"));

			IDenominationListCOMPtr strongDenomList = denomMgr->GetDenominationsStrongColl(); 

			const long count = strongDenomList->GetCount();
			_tprintf(_T("Number of Denominations: %d\n\n"), count);

			for(int i = 0; i < count; i++)
			{
				IDenominationPtr denom = strongDenomList->GetItem(i);
				com::Decimal value = denom->GetValue();
				_tprintf(_T("Denom #%d - Name = %s, Value = %s, Is Coin = %s\n"), i + 1,
					denom->GetName().GetBSTR(),
					value.ToString(),
					denom->GetIsCoin() ? _T("True") : _T("False"));
			}
		}

		_tprintf(_T("\n"));

		// Timing tests
		// --------------------------------------------------------------------------
		_tprintf(_T("+--------------+\n"));
		_tprintf(_T("| Timing Tests |\n"));		
		_tprintf(_T("+--------------+\n\n"));
		
		TimingTests(denomMgr);

		_tprintf(_T("\n"));
	}
	catch(_com_error& ex)
	{
		_tprintf(ex.ErrorMessage());
	}

	CoUninitialize();
		
	::system("PAUSE");

	return 0;
}

