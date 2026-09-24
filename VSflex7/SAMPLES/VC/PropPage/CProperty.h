// CProperty.h: declaration of the CProperty (control property) class

// This code was adapted from the ATLCON demo that ships with Microsoft 
// Visual Studio 6.
//

#ifndef __CPROPERTYDEFINED__
#define __CPROPERTYDEFINED__

class CProperty
{
public:

	// variables
	DISPID				m_dispid;
	VARTYPE				m_vtNative;
	BOOL				m_bDirty;
	CComBSTR			m_bstrDesc;
	CComBSTR			m_bstrList;
	CComVariant			m_value;
	CComPtr<IDispatch>	m_spDisp;
	CComPtr<ITypeInfo>	m_spInfo;

	// creator
	CProperty(IDispatch* pDisp, DISPID di, VARIANT val, BSTR bstrDesc, ITypeInfo* p = NULL)
	{
		m_dispid   = di;
		m_value    = val;
		m_bDirty   = false;
		m_bstrDesc = bstrDesc;
		m_vtNative = m_value.vt;
		m_spDisp   = pDisp;
		m_spInfo   = p;
		GetEnumValues();
	}

	// mark this property as dirty
	void SetDirty(BOOL bDirty = true)
	{
		m_bDirty = bDirty;
	}

	// tell whether this property is dirty
	BOOL IsDirty()
	{
		return m_bDirty;
	}

	// get string representation of property value
	HRESULT GetStringValue(BSTR* pbstr)
	{
        // see if the control wants to provide a custom string
        CComQIPtr<IPerPropertyBrowsing> spPPB(m_spDisp);
        if (spPPB != NULL) {
            if (S_OK == spPPB->GetDisplayString(m_dispid, pbstr))
                return S_OK;
        }

        // get the string ourselves
		if (!m_spInfo) {

			// booleans get translated into True/False
			if (m_vtNative == VT_BOOL) {
				CComBSTR bstr = (m_value.boolVal == VARIANT_TRUE)? L"True": L"False";
				*pbstr = bstr.Detach();
				return S_OK;
			}

			// colors are shown in hex
			if (m_vtNative == VT_UI4) {
				TCHAR szColor[20];
				wsprintf(szColor, "     %#08lx", V_I4(&m_value));
				CComBSTR bstr = szColor;
				*pbstr = bstr.Detach();
				return S_OK;
			}

			// others have names
			CComVariant v = m_value;
			v.ChangeType(VT_BSTR);
			*pbstr = v.bstrVal;
			v.bstrVal = NULL;
			return S_OK;
		}
		
		// get type attributes
		USES_CONVERSION;
		TYPEATTR* pta = NULL;
		m_spInfo->GetTypeAttr(&pta);

		// handle enumerations
		if (pta && pta->typekind == TKIND_ENUM) {
			VARDESC* pvd = NULL;
			for (int i = 0; i < pta->cVars; i++) {
				m_spInfo->GetVarDesc(i, &pvd);
				if (pvd->lpvarValue->lVal == m_value.lVal) {
					DISPID idMember = pvd->memid;
					CComBSTR bstrName;
					m_spInfo->GetDocumentation(idMember, &bstrName, NULL, NULL, NULL);
					TCHAR buf[256];
					wsprintf(buf, _T("%d - %s"), pvd->lpvarValue->lVal, OLE2T(bstrName));
					*pbstr = T2BSTR(buf);
				}
				m_spInfo->ReleaseVarDesc(pvd);
			}
		}
		
		// release type attributes
		if (pta) m_spInfo->ReleaseTypeAttr(pta);
		return S_OK;
	}

	// get category string
	HRESULT GetCategory(BSTR* pbstr)
	{
		// assume we won't get a category
		CComBSTR bstrCat;

		// get pointer to ICategorizeProperties interface
		CComPtr<ICategorizeProperties> spCat;
		m_spDisp->QueryInterface(IID_ICategorizeProperties, (void**)&spCat);

		// if we got it, get the category ID
		if (spCat != NULL) {
			PROPCAT pcat = PROPCAT_Nil;
			if (S_OK == spCat->MapPropertyToCategory(m_dispid, &pcat)) {

				// get from control
				spCat->GetCategoryName(pcat, LOCALE_USER_DEFAULT, &bstrCat);

				// if that failed, handle stock categories
				if (bstrCat.Length() == 0) {
					switch (pcat) {
						case PROPCAT_Nil:			bstrCat = "Nil";			break;
						case PROPCAT_Misc:			bstrCat = "Misc";			break;
						case PROPCAT_Font:			bstrCat = "Font";			break;
						case PROPCAT_Position:		bstrCat = "Position";		break;
						case PROPCAT_Appearance:	bstrCat = "Appearance";		break;	
						case PROPCAT_Behavior:		bstrCat = "Behavior";		break;
						case PROPCAT_Data:			bstrCat = "Data";			break;
						case PROPCAT_List:			bstrCat = "List";			break;
						case PROPCAT_Text:			bstrCat = "Text";			break;
						case PROPCAT_Scale:			bstrCat = "Scale";			break;
						case PROPCAT_DDE:			bstrCat = "DDE";			break;
					}
				}
			}
		}

		// if we have no category string, use a reasonable default
		if (bstrCat.Length() == 0) bstrCat = "Misc";

		// done
		*pbstr = bstrCat.Copy();
		return S_OK;
	}

	// set value to given string
	HRESULT SetValue(LPCTSTR pstr, IUnknown* pUnk = NULL)
	{
		// if we have a list, set value to number at the start of the string
		// (all enumerations have this format, except booleans which are handled as a special case)
		if (m_spInfo)
			return SetValue(_ttol(pstr), pUnk);

		// convert string to value
		CComVariant	vNewValue;
		switch (m_vtNative) {
			case VT_BOOL: // boolean
				vNewValue = (lstrcmpi(pstr, _T("True")) == 0)? true: false;
				break;
			case VT_UI4:  // color
				vNewValue = HexToLong(pstr);
				break;
			default:
				vNewValue = pstr;
				vNewValue.ChangeType(m_vtNative);
		}

		// get dispatch from pUnk or use internal
		CComQIPtr<IDispatch, &IID_IDispatch> spDisp(pUnk? pUnk: m_spDisp);
		CComDispatchDriver dd(spDisp);

		// put property and save new value if successful
		HRESULT hr = dd.PutProperty(m_dispid, &vNewValue);
		if (hr == S_OK) m_value = vNewValue;
		return hr;
	}

	// convert hex string into long
	long HexToLong(LPCTSTR pstr)
	{
		long val = 0;
		for (LPTSTR sz = (LPTSTR)pstr; sz && *sz; sz++) {
			TCHAR ch = tolower(*sz);
			if (ch >= '0' && ch <= '9') {
				val = val * 16 + (ch - '0');
			} else if (ch >= 'a' && ch <= 'h') {
				val = val * 16 + (ch - 'a' + 10);
			} else {
				val = 0;
			}
		}
		return val;
	}

	// set value to given enumeration value
 	HRESULT SetValue(int nEnum, IUnknown* pUnk = NULL)
	{
		// convert index into value
		CComVariant	vNewValue = nEnum;
		if (m_vtNative == VT_BOOL) vNewValue = (nEnum == 0)? true: false;

		// get dispatch from pUnk or use internal
		CComQIPtr<IDispatch, &IID_IDispatch> spDisp(pUnk? pUnk: m_spDisp);
		CComDispatchDriver dd(spDisp);

		// put property and save new value if successful
		HRESULT hr = dd.PutProperty(m_dispid, &vNewValue);
		if (hr == S_OK) m_value = vNewValue;
		return hr;
	}

	// build list of settings for this property
	HRESULT GetEnumValues()
	{
		USES_CONVERSION;
		
		// start with empty string
		m_bstrList = "";

		// no info? we can handle booleans and colors
		if (!m_spInfo) {
			switch (m_vtNative) {
				case VT_BOOL:
					m_bstrList = "True|False";
					break;
				case VT_UI4:
					m_bstrList = "...";
					break;
			}
			return S_OK;
		}

		// get list
		TYPEATTR* pta = NULL;
		m_spInfo->GetTypeAttr(&pta);
		if (pta && pta->typekind == TKIND_ENUM) {
			VARDESC* pvd = NULL;
			for (int i = 0; i < pta->cVars; i++) {
				m_spInfo->GetVarDesc(i, &pvd);
				DISPID idMember = pvd->memid;
				CComBSTR bstrName;
				m_spInfo->GetDocumentation(idMember, &bstrName, NULL, NULL, NULL);
				TCHAR buf[256];
				wsprintf(buf, _T("%d - %s"), pvd->lpvarValue->lVal, OLE2T(bstrName));
				if (m_bstrList.Length()) m_bstrList.Append("|");
				m_bstrList.Append(buf);
				m_spInfo->ReleaseVarDesc(pvd);
			}
		}
		if (pta) m_spInfo->ReleaseTypeAttr(pta);

		// all done
		return S_OK;
	}
};

#endif
