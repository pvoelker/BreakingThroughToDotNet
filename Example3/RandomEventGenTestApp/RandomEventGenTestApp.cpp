// Breaking Through to .NET - Example 3 - COM interop event mechanism
// March 2014
// Paul Voelker

#include "stdafx.h"
#include "RandomEventSink.h"

#import "..\TypeLibs\RandomEventGenerator.tlb" no_namespace

HANDLE m_ThreadExit;

// ----------------------------------------------------------------------------
// Event callback processing thread
// ----------------------------------------------------------------------------
DWORD ThreadProc(LPVOID lpdwThreadParam) 
{
	::CoInitializeEx(0, COINIT_MULTITHREADED);

	try
	{
		IRandomEventGeneratorCOMPtr randomEventGen(__uuidof(RandomEventGeneratorCOM));

		RandomEventGeneratorEventSink eventSink;

		randomEventGen->Subscribe((IRandomEventCOMEvents*)(&eventSink));
		
		DWORD waitForResult = WAIT_FAILED;
		while(waitForResult != WAIT_OBJECT_0)
		{
			waitForResult = ::WaitForSingleObject(m_ThreadExit, INFINITE);
		}

		randomEventGen->Unsubscribe();
	}
	catch(_com_error& ex)
	{
		_tprintf(ex.ErrorMessage());
		return 0;
	}

	::CoUninitialize();
	
	return 0;
}

// ----------------------------------------------------------------------------
// Application entry point
// ----------------------------------------------------------------------------
int _tmain(int argc, _TCHAR* argv[])
{
	int runTimeSeconds = 10;

	m_ThreadExit = ::CreateEvent(NULL, TRUE, FALSE, _T("Thread Exit"));

	DWORD threadId;
	if (::CreateThread(NULL, 0, (LPTHREAD_START_ROUTINE)&ThreadProc,
		(LPVOID)NULL, 0, &threadId) == NULL)
	{
		_tprintf(_T("Error creating thread\n"));
		return(1);
	}
	else
	{
		_tprintf(_T("Event processing thread created, waiting %d seconds ")
			_T("for random events...\n\n"), runTimeSeconds);
		::Sleep(runTimeSeconds * 1000);
		::SetEvent(m_ThreadExit);
	}

	_tprintf(_T("\nTime's up!!!!\n\n"));

	::system("PAUSE");

	return 0;
}


