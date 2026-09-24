/////////////////////////////////////////////////////////////////////////////
// FlexDSLinks - implements an IVSFlexGridDataSource interface for
//               providing link data
//

#ifndef _FLEXDSLINKSDECLARED
#define _FLEXDSLINKSDECLARED

class CFlexDSLinks
	: public IVSFlexDataSource
{
public:
	ULONG m_cRef;

// constructor
	CFlexDSLinks() { m_cRef = 0; }

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
		*pFields = 9;
		return S_OK;
	}
	STDMETHOD(raw_GetRecordCount)(long * pRecords)
	{
		*pRecords = m_pData->GetLinkCount();
		return S_OK;
	}
    STDMETHOD(raw_GetFieldName)(long Field, BSTR * pFieldName)
	{
		_bstr_t str;
		switch (Field)
		{
			case 0: str = "#";			break;
			case 1: str = "ID";			break;
			case 2: str = "From";		break;
			case 3: str = "To";			break;
			case 4: str = "Length";		break;
			case 5: str = "Section";	break;
			case 6: str = "Nu";			break;
			case 7: str = "Nv";			break;
			case 8: str = "Nw";			break;
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
			// most fields come directly from the data
			case 0: var = Record;								break;
			case 1: var = m_pData->m_Links[Record].szName;		break;
			case 2: var = (long)m_pData->m_Links[Record].i;		break;
			case 3: var = (long)m_pData->m_Links[Record].j;		break;
			case 5: var = (long)m_pData->m_Links[Record].iSec;	break;
			case 6: var = m_pData->m_Links[Record].u;			break;
			case 7: var = m_pData->m_Links[Record].v;			break;
			case 8: var = m_pData->m_Links[Record].w;			break;
                
            // field 4 is calculated: get the link length
            case 4:
				{
					int iFrom = m_pData->m_Links[Record].i;
					int iTo   = m_pData->m_Links[Record].j;
					double dx = m_pData->m_Nodes[iFrom].x - m_pData->m_Nodes[iTo].x;
					double dy = m_pData->m_Nodes[iFrom].y - m_pData->m_Nodes[iTo].y;
					double dz = m_pData->m_Nodes[iFrom].z - m_pData->m_Nodes[iTo].z;
					var = sqrt(dx*dx + dy*dy + dz*dz);
				}
				break;
		}

		// done
		var.ChangeType(VT_BSTR);
		_bstr_t str = V_BSTR(&var);
		*pVal = str.copy();
		return S_OK;
	}
	STDMETHOD(raw_SetData)(long Field, long Record, BSTR newData)
	{
		// sanity
		ATLASSERT(Record >= 0 && Record < m_pData->GetLinkCount() && Field >= 0 && Field < 9);

		// convert new value to LPTSTR
		USES_CONVERSION;
		LPTSTR sz = W2T(newData);

		// set data
		int i;
		TCHAR szMsg[100];
		switch (Field)
		{
			// setting link name
			case 1:

				// make sure link name is unique
				for (i = 0; i < m_pData->GetLinkCount(); i++)
				{
					if (i != Record && _tcscmp(sz, m_pData->m_Links[i].szName) == 0)
					{
						wsprintf(szMsg, "Sorry, link names must be unique.\r\nThe name '%s' already belongs to link #%d.", sz, i);
						MessageBox(NULL, szMsg, "Invalid Link Name", MB_ICONWARNING);
						return S_OK;
					}
				}

				// name is unique, so save it
				_tcsncpy(m_pData->m_Links[Record].szName, sz, 29);
				return S_OK;
        
	        // make sure connecting nodes different and in the valid range
		    case 2:
			    i = atoi(sz);
				if (i < 0 || i >= m_pData->GetNodeCount())
				{
					wsprintf(szMsg, "Valid node numbers are between 0 asnd %d.", m_pData->GetNodeCount() - 1);
					MessageBox(NULL, szMsg, "Invalid Node Number", MB_ICONWARNING);
					return S_OK;
				}
				if (i == m_pData->m_Links[Record].j)
				{
					MessageBox(NULL, "Links must connect two different nodes.", "Invalid Node Number", MB_ICONWARNING);
					return S_OK;
				}
	            m_pData->m_Links[Record].i = i;
				return S_OK;

			case 3:
			    i = atoi(sz);
				if (i < 0 || i >= m_pData->GetNodeCount())
				{
					wsprintf(szMsg, "Valid node numbers are between 0 asnd %d.", m_pData->GetNodeCount() - 1);
					MessageBox(NULL, szMsg, "Invalid Node Number", MB_ICONWARNING);
					return S_OK;
				}
				if (i == m_pData->m_Links[Record].i)
				{
					MessageBox(NULL, "Links must connect two different nodes.", "Invalid Node Number", MB_ICONWARNING);
					return S_OK;
				}
	            m_pData->m_Links[Record].j = i;
				return S_OK;

	        // section number
			case 5:
			    i = atoi(sz);
				if (i < 1000 || i > 2000)
				{
					MessageBox(NULL, "Valid section numbers are between 1000 and 2000.", "Invalid Section Number", MB_ICONWARNING);
					return S_OK;
				}
				m_pData->m_Links[Record].iSec = i;
				return S_OK;
            
			// normal vector needs no validation
			case 6: m_pData->m_Links[Record].u = atof(sz); break;
			case 7: m_pData->m_Links[Record].v = atof(sz); break;
			case 8: m_pData->m_Links[Record].w = atof(sz); break;
		}        
		return S_OK;
	}

// Pointer to the actual data
	CTheData* m_pData;
};

#endif // _FLEXDSLINKSDECLARED
