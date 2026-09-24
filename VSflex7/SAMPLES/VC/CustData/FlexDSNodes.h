/////////////////////////////////////////////////////////////////////////////
// FlexDSNodes - implements an IVSFlexGridDataSource interface for
//               providing node data
//

#ifndef _FLEXDSNODESDECLARED
#define _FLEXDSNODESDECLARED

class CFlexDSNodes
	: public IVSFlexDataSource
{
public:
	ULONG m_cRef;

// constructor
	CFlexDSNodes() { m_cRef = 0; }

// IUnknown members
// we implement IUnknown and IVSFlexDataSource
	STDMETHOD(QueryInterface)(REFIID riid, LPVOID *ppv) {
		if (riid == IID_IUnknown || riid == __uuidof(IVSFlexDataSource)) {
			*ppv = (LPVOID)this;
			((LPUNKNOWN)*ppv)->AddRef();
			return S_OK;
		}
		*ppv = NULL;
		return E_NOINTERFACE;
	}
    STDMETHODIMP_(ULONG) AddRef()	{ return ++m_cRef; }
    STDMETHODIMP_(ULONG) Release()	{ return --m_cRef; } // we're a static class, so don't delete anything

// IVSFlexDataSource members
// Wrapper methods for error-handling
	long GetFieldCount()
	{
		long nFields;
		raw_GetFieldCount(&nFields);
		return nFields;
	}
	long GetRecordCount()
	{
		long nRecords;
		raw_GetRecordCount(&nRecords);
		return nRecords;
	}
	_bstr_t GetFieldName(long Field)
	{
		_bstr_t str;
		raw_GetFieldName(Field, (BSTR*)&str);
		return str;
	}
	_bstr_t GetData(long Field, long Record)
	{
		_bstr_t str;
		raw_GetData(Field, Record, (BSTR*)&str);
		return str;
	}
    HRESULT SetData(long Field, long Record, _bstr_t newData)
	{
		return raw_SetData(Field, Record, newData);
	}

// Raw methods provided by interface
	STDMETHOD(raw_GetFieldCount)(long * pFields)
	{
		*pFields = 5;
		return S_OK;
	}
	STDMETHOD(raw_GetRecordCount)(long * pRecords)
	{
		*pRecords = m_pData->GetNodeCount();
		return S_OK;
	}
    STDMETHOD(raw_GetFieldName)(long Field, BSTR * pFieldName)
	{
		_bstr_t str;
		switch (Field)
		{
			case 0: str = "#";			break;
			case 1: str = "ID";	break;
			case 2: str = "x";	break;
			case 3: str = "y";	break;
			case 4: str = "z";	break;
		}
		*pFieldName = str.copy();
		return S_OK;
	}
	STDMETHOD(raw_GetData)(long Field, long Record, BSTR * pVal)
	{
		// get data
		_variant_t var;
        switch (Field)
		{
			case 0: var = Record;							break;
			case 1: var = m_pData->m_Nodes[Record].szName;	break;
			case 2: var = (long)m_pData->m_Nodes[Record].x;	break;
			case 3: var = (long)m_pData->m_Nodes[Record].y;	break;
			case 4: var = (long)m_pData->m_Nodes[Record].z;	break;
		}

		// done
		var.ChangeType(VT_BSTR);
		_bstr_t str = V_BSTR(&var);
		*pVal = str.copy();
		return S_OK;
	}
	STDMETHOD(raw_SetData)(long Field, long Record, BSTR newVal)
	{
		// sanity
		ATLASSERT(Record >= 0 && Record < m_pData->GetNodeCount() && Field >= 0 && Field < 6);

		// convert new value to LPTSTR
		USES_CONVERSION;
		LPTSTR sz = W2T(newVal);

		// set data
		int i;
		TCHAR szMsg[100];
		switch (Field)
		{
			// setting node name
			case 1:

				// make sure node name is unique
				for (i = 0; i < m_pData->GetNodeCount(); i++)
				{
					if (i != Record && _tcscmp(sz, m_pData->m_Nodes[i].szName) == 0)
					{
						wsprintf(szMsg, "Sorry, node names must be unique.\r\nThe name '%s' already belongs to node #%d.", sz, i);
						MessageBox(NULL, szMsg, "Invalid Node Name", MB_ICONWARNING);
						return S_OK;
					}
				}

				// name is unique, so save it
				_tcsncpy(m_pData->m_Nodes[Record].szName, sz, 29);
				return S_OK;
        
			// coordinates need no validation
			case 2: m_pData->m_Nodes[Record].x = atof(sz); break;
			case 3: m_pData->m_Nodes[Record].y = atof(sz); break;
			case 4: m_pData->m_Nodes[Record].z = atof(sz); break;
		}        
		return S_OK;
	}

// Pointer to the actual data
	CTheData* m_pData;
};

#endif // _FLEXDSNODESDECLARED
