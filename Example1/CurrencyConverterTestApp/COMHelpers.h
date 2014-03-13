// Re-usable COM Helper Functions
// Paul Voelker

#include "stdafx.h"

#include <comdef.h>
#include <comdefsp.h>

CString GetCOMError(IErrorInfo* pErrorInfo)
{
	CString retVal = _T("No error");

	if(pErrorInfo)
	{
		// Strings in raw COM interfaces are passed as BSTRs.  Memory for BSTRs
		// must be allocated and released.
		// See: http://msdn.microsoft.com/en-us/library/xda6xzx7(v=vs.110).aspx
		BSTR errDesc;
		pErrorInfo->GetDescription(&errDesc);
		retVal = errDesc;
		::SysFreeString(errDesc);
	}

	return retVal;
}

CString GetLastCOMError()
{
	IErrorInfoPtr spErrorInfo;
	::GetErrorInfo(0, &spErrorInfo);
	return GetCOMError(spErrorInfo);
}
