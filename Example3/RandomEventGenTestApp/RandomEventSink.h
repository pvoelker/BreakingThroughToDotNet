#pragma once

#include "stdafx.h"

#import "..\TypeLibs\RandomEventGenerator.tlb" no_namespace

// Need to implement a COM object in order to support COM callback mechanism
// Reference: http://msdn.microsoft.com/en-us/library/5hhehwba(v=vs.110).aspx
class RandomEventGeneratorEventSink : IRandomEventCOMEvents
{
private:
	// BEGIN: IUnknown implementation
	// ------------------------------------------------------------------------
	LONG m_ReferenceCount;
	// ------------------------------------------------------------------------
	// END: IUnknown implementation

public:
	RandomEventGeneratorEventSink()
	{
	}

    HRESULT __stdcall raw_TimerFired()
	{
		CTime currentTime = CTime::GetCurrentTime();
		_tprintf(_T("Timer fired at: ") + currentTime.Format(_T("%c")) + _T("\n"));

		return 0;
	}

	// BEGIN: IUnknown implementation
	// ------------------------------------------------------------------------
	HRESULT APIENTRY QueryInterface(REFIID iid, void** ppvObject)
	{
		// Match the interface and return the proper pointer
		if (iid == IID_IUnknown)
		{
			*ppvObject = dynamic_cast<IUnknown*>(this);
		}
		else if (iid == __uuidof(IRandomEventCOMEvents))
		{
			*ppvObject = dynamic_cast<IRandomEventCOMEvents*>(this);
		}
		else
		{
			// No matching interfaces found
			*ppvObject = NULL;
			return E_NOINTERFACE;
		}

		this->AddRef();

		return S_OK;
	}

	ULONG APIENTRY AddRef() { return ++m_ReferenceCount; }

	ULONG APIENTRY Release() { return --m_ReferenceCount; }	
	// ------------------------------------------------------------------------
	// END: IUnknown implementation
};
