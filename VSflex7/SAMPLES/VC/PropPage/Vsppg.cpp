// VSPPG.cpp : Implementation of CVSPPG
#include "stdafx.h"
#include "VSPPG7.h"
#include "VSPPG.h"


/////////////////////////////////////////////////////////////////////////////
// CVSPPG: interface implementation

// prevent ESC key from dismissing dialog while we're editing
HRESULT CVSPPG::TranslateAccelerator(MSG *pMsg)
{
	// disable VK_ESCAPE key while editing grid values
	if (pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		// if editing grid, pass the message to the grid and hide the message from the dialog
		CWindow wnd = (HWND)m_spGrid->EditWindow;
		if (wnd != NULL)
		{
			wnd.SendMessage(WM_KEYDOWN, VK_ESCAPE);
			pMsg->wParam = 0;
			return S_OK;
		}
	}

	// use default ATL implementation
	return IPropertyPageImpl<CVSPPG>::TranslateAccelerator(pMsg);
}

// apply changes to controls
HRESULT CVSPPG::Apply(void)
{
	// scan properties
	long iRow, iCnt;
	iCnt = (m_spGrid)? m_spGrid->Rows: 0; // sanity
	for (iRow = 1; iRow < iCnt; iRow++) {

		// get property from grid row
		_variant_t vProp = m_spGrid->GetRowData(iRow);
		if (vProp.vt != VT_I4) continue;
		CProperty* pProp = (CProperty*)V_I4(&vProp);

		// if not dirty, don't waste time
		if (!pProp->IsDirty()) continue;
		pProp->SetDirty(false);

		// apply changes to all controls
		USES_CONVERSION;
		_bstr_t bstrName = m_spGrid->GetTextMatrix(iRow, 2);
		_bstr_t bstrVal  = m_spGrid->GetTextMatrix(iRow, 3);
		for (UINT i = 0; i < m_nObjects; i++) {
			if (S_OK != pProp->SetValue(bstrVal, m_ppUnk[i])) {
				TCHAR buf[256];
				wsprintf(buf, _T("Cannot set property '%s' to value '%s'"), OLE2T(bstrName), OLE2T(bstrVal)),
				MessageBox(buf, _T("Bad Property Value"), MB_ICONERROR);
			}
		}

		// make sure we're displaying the right value
		CComBSTR bstrFresh;
		pProp->GetStringValue(&bstrFresh);
		m_spGrid->PutTextMatrix(iRow, 3, _bstr_t(bstrFresh));

		// make it non-bold
		m_spGrid->PutCell(flexcpFontBold, iRow, 3L, vtMissing, vtMissing, (long)false);
	}

	// we're clean and done
	SetDirty(false);
	return S_OK;
}


/////////////////////////////////////////////////////////////////////////////
// CVSPPG: message handlers

// initialize dialog
LRESULT CVSPPG::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	// initialize font
	static FONTDESC _fdDefault = { sizeof(FONTDESC), OLESTR("MS Sans Serif"), FONTSIZE(8),
								   FW_NORMAL, DEFAULT_CHARSET, false, false, false };
	OleCreateFontIndirect(&_fdDefault, IID_IFont, (void**)&m_spFont);
	CComQIPtr<IFontDisp, &IID_IFontDisp> pFont(m_spFont);
	HFONT hFont;
	m_spFont->get_hFont(&hFont);

	// calculate font height to dimension controls
	int nFontHeight = GetFontHeight();

	// initialize "by categories" button
	RECT rc, rcBtn;
	CWindow wndByCat = GetDlgItem(IDC_BYCAT);
    wndByCat.GetClientRect(&rcBtn);
	GetClientRect(&rc);
	InflateRect(&rc, -BORDER, -BORDER);
	rc.top    = rc.bottom - nFontHeight * 2;
	rc.right  = rc.left + rcBtn.right;
	wndByCat.MoveWindow(&rc);
	m_hIcon = (HICON)LoadImage(_Module.GetResourceInstance(), MAKEINTRESOURCE(IDI_BYCAT), IMAGE_ICON, 0, 0, 0);
	wndByCat.SendMessage(BM_SETIMAGE, IMAGE_ICON, (LPARAM)m_hIcon);

	// initialize help window
	GetClientRect(&rc);
	InflateRect(&rc, -BORDER, -BORDER);
	rc.top  = rc.bottom - nFontHeight * 2;
	rc.left = rc.left + rcBtn.right + BORDER;
	m_wndHelp.Attach(GetDlgItem(IDC_HELPTEXT));
	m_wndHelp.MoveWindow(&rc);
	m_wndHelp.SendMessage(WM_SETFONT, (WPARAM)hFont);

	// create grid control
	if (S_OK != m_spGrid.CreateInstance(__uuidof(VSFlexGrid))) {
		m_wndHelp.SetWindowText(MSG_NOFLEX);
		return -1;
	}

	// initialize ATL hosting
	AtlAxWinInit();

	// create ActiveX control host window
	GetClientRect(&rc);
	InflateRect(&rc, -BORDER, -BORDER);
	rc.bottom -= (nFontHeight * 2 + BORDER);
	m_wndHost.Create(m_hWnd, rc, _T(""), STYLE_HOST, 0, IDC_GRID);

	// attach FlexGrid control to host window
	CComPtr<IAxWinHostWindow> spHost;
	m_wndHost.QueryHost(&spHost);
	spHost->AttachControl(m_spGrid, m_wndHost);

	// connect events
	AtlAdviseSinkMap(this, true);

	// initialize grid control
	m_spGrid->PutFont(pFont);
	InitGrid();
	
	// populate grid control
	FillPropertyList();

	// initialize display
	if (m_bHasCat) {
		wndByCat.SendMessage(BM_SETCHECK, true);
		DisplayCategorized();
	} else {
		wndByCat.EnableWindow(false);
		DisplayAlphabetic();
	}

	// no double clicks yet
	m_bDblClick = false;

	// all done
	m_wndHost.ShowWindow(SW_SHOW);
    m_wndHost.SetFocus();
	return 0;
}

// done
LRESULT CVSPPG::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	// release grid control
	if (m_spGrid) m_spGrid.Release();

	// destroy button icon
	if (m_hIcon) DestroyIcon(m_hIcon);

	// allow default processing
	bHandled = false;
	return 0;
}

// switch views
LRESULT CVSPPG::OnClickedByCat(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	CWindow wnd = hWndCtl;
	if (wnd.SendMessage(BM_GETCHECK)) {
		DisplayCategorized();
	} else {
		DisplayAlphabetic();
	}
	return 0;
}


/////////////////////////////////////////////////////////////////////////////
// CVSPPG: internal methods

void CVSPPG::DisplayCategorized()
{
	// if we can't display categorized, display alphabetic instead
	if (!m_bHasCat) {
		DisplayAlphabetic();
		return;
	}

	// make sure we have a control
	if (m_spGrid == NULL) return;

	// freeze to avoid flicker
    m_spGrid->Redraw = flexRDNone;							

	// remove any existing subtotals (groups)
    m_spGrid->Subtotal(flexSTClear);						

	// sort by category/property name
    m_spGrid->Select(1, 1, 1L, 2L); 
    m_spGrid->Sort = flexSortStringAscending;

	// add groups on column 1 (category), with a back color and bold font
    m_spGrid->Subtotal(flexSTNone, 1L, vtMissing, vtMissing,
					   (long)m_spGrid->GridColor, vtMissing, (long)true);

	// show outline column
    m_spGrid->PutColHidden(0, false);						

	// autosize on property names
    m_spGrid->AutoSize(2, vtMissing, vtMissing, 300L);		

	// to look nice
    m_spGrid->GridLines = flexGridFlatVert;					

	// reset display
    m_spGrid->TopRow = 1;									
    m_spGrid->Select(2, m_spGrid->Cols - 1);			
    m_spGrid->Redraw = flexRDBuffered;
}

void CVSPPG::DisplayAlphabetic()
{
	// make sure we have a control
	if (m_spGrid == NULL) return;

	// freeze to avoid flicker
    m_spGrid->Redraw = flexRDNone;							

	// remove any existing subtotals (groups)
    m_spGrid->Subtotal(flexSTClear);						

	// sort by property name
    m_spGrid->Col = 2;										
    m_spGrid->Sort = flexSortStringAscending;

	// hide outline column
    m_spGrid->PutColHidden(0, true);						

	// autosize on property names
    m_spGrid->AutoSize(2, vtMissing, vtMissing, 300L);		

    // to look nice
	m_spGrid->GridLines = flexGridFlat;						

	// reset display
    m_spGrid->TopRow = 1;									
    m_spGrid->Select(1, m_spGrid->Cols - 1);
    m_spGrid->Redraw = flexRDBuffered;
}

void CVSPPG::InitGrid()
{
	if (m_spGrid == NULL) return;							// sanity
    m_spGrid->Rows = 1;								        // start empty
    m_spGrid->Cols = 4;										// outline, category, property, value
    m_spGrid->PutTextMatrix(0, 2, L"Property");				// column titles
    m_spGrid->PutTextMatrix(0, 3, L"Value");				// column titles
    m_spGrid->Editable = flexEDKbd;							// editable, handle double-clicks ourselves
    m_spGrid->OutlineCol = 0;								// set outline column properties
    m_spGrid->PutColWidth(0, 230);							// narrow outline column
    m_spGrid->PutColHidden(1, true);						// hide categories (they'll just spill on group rows)
    m_spGrid->OutlineBar = flexOutlineBarSymbolsLeaf;		// no tree, just symbols
    m_spGrid->MergeCells = flexMergeSpill;					// allow categories to spill into property column
    m_spGrid->PutColAlignment(-1, flexAlignLeftTop);		// align all columns to the left
    m_spGrid->AllowSelection = false;						// select a single cell at a time
    m_spGrid->AllowUserResizing = flexResizeColumns;		// give user freedom
    m_spGrid->ScrollTrack = true;							// scroll as the user drags the scroll thumb
	m_spGrid->OwnerDraw = flexODOver;						// use ownerdraw to show colors
    m_spGrid->FixedCols = 0;								// to look nice
    m_spGrid->ExtendLastCol = true;
	m_spGrid->Ellipsis = flexEllipsisEnd;
    m_spGrid->BackColorBkg = m_spGrid->GridColor;
    m_spGrid->HighLight = flexHighlightNever;
}

// measure font height
int CVSPPG::GetFontHeight()
{
	TEXTMETRIC tm;
	HFONT      hFont;
	HDC        hdc = GetDC();
	m_spFont->get_hFont(&hFont);
	HFONT hOldFont = (HFONT)SelectObject(hdc, hFont);
	GetTextMetrics(hdc, &tm);
	SelectObject(hdc, hOldFont);
	ReleaseDC(hdc);
	return tm.tmHeight + tm.tmExternalLeading;
}

// helper
HRESULT CVSPPG::GetEnumTypeInfo(ITypeInfo *pTI, HREFTYPE hrt, ITypeInfo** ppEnumInfo)
{
	// get type info
	CComPtr<ITypeInfo> spTypeInfo;
	HRESULT hr = E_FAIL;
	hr = pTI->GetRefTypeInfo(hrt, &spTypeInfo);
	if (FAILED(hr)) return hr;

	// get type attributes
	TYPEATTR *pta = NULL;
	spTypeInfo->GetTypeAttr(&pta);
	if (!pta) return E_FAIL;

	// get info
	hr = E_FAIL;
	switch (pta->typekind) {
		case TKIND_ALIAS:
			if (pta->tdescAlias.vt == VT_USERDEFINED) {
				GetEnumTypeInfo(spTypeInfo,pta->tdescAlias.hreftype, ppEnumInfo);
				if (*ppEnumInfo) hr = S_OK;
			} else if (pta->tdescAlias.vt == VT_UI4) {
				hr = S_OK;
			}
			break;
		case TKIND_ENUM:
			spTypeInfo.CopyTo(ppEnumInfo);
			if (*ppEnumInfo) hr = S_OK;
			break;
	}
	
	// all done
	spTypeInfo->ReleaseTypeAttr(pta);
	return hr;
}

// scan control for its properties, the fill property list
void CVSPPG::FillPropertyList()
{
	// assume we can't categorize properties
	m_bHasCat = false;

	// sanity
	if (m_spGrid == NULL) return;

	// clear any existing contents
	ClearPropertyList();

	// get dispatch interface for first target control
	if (!m_ppUnk[0]) return;
	CComQIPtr<IDispatch, &IID_IDispatch> spDisp(m_ppUnk[0]);
	if (!spDisp) return;

	// see if we can categorize properties
	CComPtr<ICategorizeProperties> spCat;
	spDisp->QueryInterface(IID_ICategorizeProperties, (void**)&spCat);
	if (spCat) m_bHasCat = true;

	// get type info
	CComPtr<ITypeInfo> spTypeInfo;
	spDisp->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &spTypeInfo);
	if (!spTypeInfo) return;

	// get type attributes
	TYPEATTR* pta;
	spTypeInfo->GetTypeAttr(&pta);
	int nProps = pta->cFuncs;

	// if type is interface, get dual
	if (pta->typekind == TKIND_INTERFACE) {
		nProps = 0;
		HREFTYPE hRef;
		CComPtr<ITypeInfo> spInfoTemp;
		if (S_OK == spTypeInfo->GetRefTypeOfImplType(-1, &hRef) &&
			S_OK == spTypeInfo->GetRefTypeInfo(hRef, &spInfoTemp)) {
			spTypeInfo->ReleaseTypeAttr(pta);
			spTypeInfo = spInfoTemp;
			spTypeInfo->GetTypeAttr(&pta);
			nProps = pta->cFuncs;
		}
	}

	// release type attributes
	spTypeInfo->ReleaseTypeAttr(pta);
	pta = NULL;

	// add each property
	for (int i = 0; i < nProps; i++) {

		// get function description:
		// skip all that are not property get, restricted, and hidden
		FUNCDESC* pfd;
		spTypeInfo->GetFuncDesc(i, &pfd);
		
		// skip unwanted stuff
		BOOL bSkip = false;
		if ((pfd->invkind & DISPATCH_PROPERTYGET) == 0) bSkip = true;
		if (pfd->wFuncFlags & (FUNCFLAG_FRESTRICTED | FUNCFLAG_FHIDDEN)) bSkip = true;
		if (pfd->memid == DISPID_HWND) bSkip = true;
		if (bSkip) {
			spTypeInfo->ReleaseFuncDesc(pfd);
			continue;
		}

		// handle known types
		CComPtr<ITypeInfo> spUserTypeInfo;
		CComBSTR	bstrCat, bstrName, bstrVal, bstrDoc;
		HREFTYPE	hrt;
		VARTYPE		vt;
		HRESULT		hr;
		CComVariant	varVal;
		switch (pfd->elemdescFunc.tdesc.vt) {

			// handle user-defined types
			case VT_USERDEFINED:
				hrt = pfd->elemdescFunc.tdesc.hreftype;
				vt = VT_USERDEFINED;
				hr = E_FAIL;
				hr = GetEnumTypeInfo(spTypeInfo, hrt, &spUserTypeInfo);
				if (FAILED(hr)) break;
				// FALL THROUGH

			// handle built-in types
			case VT_EMPTY:	case VT_NULL:
			case VT_BSTR:	case VT_DATE:	case VT_ERROR:	case VT_BOOL:	case VT_VARIANT:
			case VT_UINT:	case VT_I1:		case VT_UI1:	case VT_UI2:	case VT_UI4:
			case VT_INT:	case VT_I2:		case VT_I4:		case VT_R4:		case VT_R8:
			case VT_CY:		case VT_DECIMAL: {

				// get documentation
				hr = spTypeInfo->GetDocumentation(pfd->memid, &bstrName, &bstrDoc, NULL, NULL);
				if (hr != S_OK) break;
				CComDispatchDriver dd(spDisp);

				// get property value
				hr = dd.GetProperty(pfd->memid, &varVal);
				if (hr != S_OK || V_VT(&varVal) == VT_NULL) break;
				CProperty* pProp = new CProperty(spDisp, pfd->memid, varVal, bstrDoc, spUserTypeInfo);
				pProp->GetStringValue(&bstrVal);
				pProp->GetCategory(&bstrCat);

				// add property to list
				int iPos = m_spGrid->Rows;
				m_spGrid->Rows = iPos + 1;
				m_spGrid->PutTextMatrix(iPos, 1, _bstr_t(bstrCat));
				m_spGrid->PutTextMatrix(iPos, 2, _bstr_t(bstrName));
				m_spGrid->PutTextMatrix(iPos, 3, _bstr_t(bstrVal));

				// store pointer to property
				_variant_t vProp((long)pProp);
				m_spGrid->PutRowData(iPos, &vProp);
			}
		}

		// done with this function
		spTypeInfo->ReleaseFuncDesc(pfd);
		pfd = NULL;
	}

	// all done
	return;
}

// clean up
void CVSPPG::ClearPropertyList()
{
	// sanity
	if (m_spGrid == NULL) return;

	// delete properties
	long nCount = m_spGrid->Rows;
	for (long iRow = 1; iRow < nCount; iRow++) {
		_variant_t vProp = m_spGrid->GetRowData(iRow);
		if (vProp.vt == VT_I4) 
			delete (CProperty*)V_I4(&vProp);
	}
	
	// delete rows	
	m_spGrid->PutRows(1);
}

LPTSTR CVSPPG::InStr(LPTSTR str1, LPTSTR str2)
{
    if (!str1 || !*str1 || !str2 || !*str2) return NULL;
    int l = lstrlen(str2);
    for ( ; *str1; str1++)
        if (!_tcsnicmp(str1, str2, l)) return str1;
    return NULL;
}

// select next item for a cell with a ComboList, return true if successful
bool CVSPPG::SelectNextItem(long Row, long Col)
{
	// get current list, trim combo pipe if any
	USES_CONVERSION;
	LPTSTR szList = W2T(m_spGrid->ComboList);
	if (!szList || !lstrlen(szList)) return false;
	if (*szList == '|') szList++;

    // look for current text in list, fail if not found
    LPTSTR szNext = InStr(szList, W2T(m_spGrid->GetTextMatrix(Row, Col)));
    if (!szNext) return false;
    
    // look for next choice
	szNext = _tcschr(szNext, '|');
	szNext = (szNext)? szNext + 1: szList;
    
    // trim excess
    LPTSTR szEnd = _tcschr(szNext, '|');
	if (szEnd) *szEnd = 0;
    
    // set new entry
    m_spGrid->PutTextMatrix(Row, Col, szNext);
	return true;
}


/////////////////////////////////////////////////////////////////////////////
// CVSPPG: event handlers

// RowColChange: update help text 
void _stdcall CVSPPG::OnRowColChange()
{
	// AutoSearch is on for column 0 only
	m_spGrid->PutAutoSearch((m_spGrid->GetCol() == 0)
		? flexSearchFromCursor
		: flexSearchNone);

	// update description text
	USES_CONVERSION;
	_variant_t vProp = m_spGrid->GetRowData(m_spGrid->Row);
	if (vProp.vt == VT_I4 && V_I4(&vProp)) {
		CProperty* pProp = (CProperty*)V_I4(&vProp);
		LPCTSTR lpszDesc = (pProp == NULL)? _T(""): OLE2T(pProp->m_bstrDesc);
		m_wndHelp.SetWindowText(lpszDesc);
	}
}

// BeforeRowColChange: only the value column can be selected
void _stdcall CVSPPG::OnBeforeRowColChange(long OldRow, long OldCol, long NewRow, long NewCol, VARIANT_BOOL* Cancel)
{
	if (m_spGrid->Redraw && NewCol != m_spGrid->Cols - 1) {
		*Cancel = true;
		m_spGrid->Select(NewRow, m_spGrid->Cols - 1);
	}
}

// AfterEdit: mark property and page as dirty
void _stdcall CVSPPG::OnAfterEdit(long Row, long Col)
{
	// get property pointer
	_variant_t vProp = m_spGrid->GetRowData(Row);
	if (vProp.vt != VT_I4) return;
	CProperty* pProp = (CProperty*)V_I4(&vProp);

	// mark property as dirty
	pProp->SetDirty(true);
	m_spGrid->PutCell(flexcpFontBold, Row, 3L, vtMissing, vtMissing, (long)true);

	// mark page as dirty
	SetDirty(true);
}

// BeforeEdit: allow editing property value only
void _stdcall CVSPPG::OnBeforeEdit(long Row, long Col, VARIANT_BOOL* Cancel)
{
    // we can't edit total rows or label columns
    if (m_spGrid->GetIsSubtotal(Row) || Col != m_spGrid->Cols - 1) {
        *Cancel = true;
		return;
	}

	// set ComboList to value obtained from property
	_bstr_t bstrList = "";
	_variant_t vProp = m_spGrid->GetRowData(m_spGrid->Row);
	if (vProp.vt == VT_I4) {
		CProperty* pProp = (CProperty*)V_I4(&vProp);
		bstrList = pProp->m_bstrList;
	}
	m_spGrid->ComboList = bstrList;

    // use automatic double-click for editing text, manual for lists
	m_spGrid->Editable = (bstrList.length())? flexEDKbd: flexEDKbdMouse;
}

// StartEdit: if editing lists with double-click, cycle list values
void _stdcall CVSPPG::OnStartEdit(long Row, long Col, VARIANT_BOOL* Cancel)
{
    // if this is a list, double-clicking selects the next item
    if (m_bDblClick && SelectNextItem(Row, Col)) {
		*Cancel = true;
		OnAfterEdit(Row, Col);
	}

	// done
	m_bDblClick = false;
}

// DblClick: double-clicking on a group collapses/expands it
void _stdcall CVSPPG::OnDblClick()
{
	// if row is subtotal, expand/collapse it
    long r = m_spGrid->MouseRow;
    if (m_spGrid->GetIsSubtotal(r)) {
		m_bDblClick = false;
		CollapsedSettings cs = (m_spGrid->GetIsCollapsed(r) == flexOutlineCollapsed)
			? flexOutlineExpanded
			: flexOutlineCollapsed;
		m_spGrid->PutIsCollapsed(r, cs);

	// otherwise, start editing
	} else {
		m_bDblClick = true;
		m_spGrid->EditCell();
	}
}

// KeyDown: provide outlining and navigation interface
void _stdcall CVSPPG::OnKeyDown(short* KeyCode, short Shift)
{
    // special handling for cursor keys
    switch (*KeyCode) {

        // collapse/expand with cursor keys
        case 37: /* vbKeyLeft */
		case 36: /* vbKeyHome */
            if (m_spGrid->GetIsSubtotal(m_spGrid->Row))
				m_spGrid->PutIsCollapsed(m_spGrid->Row, flexOutlineCollapsed);
            if (m_spGrid->Col != m_spGrid->Cols - 1)
				m_spGrid->Col = m_spGrid->Cols - 1;
            *KeyCode = 0;
			return;

        case 39: /* vbKeyRight */
		case 35: /* vbKeyEnd */
            if (m_spGrid->GetIsSubtotal(m_spGrid->Row))
				m_spGrid->PutIsCollapsed(m_spGrid->Row, flexOutlineExpanded);
            if (m_spGrid->Col != m_spGrid->Cols - 1)
				m_spGrid->Col = m_spGrid->Cols - 1;
            *KeyCode = 0;
			return;
	}

    // when pushing control+ASCII, look for property
    if (Shift < 2 || *KeyCode < (short)'A' || *KeyCode > (short)'Z') return;
            
    // look from current row down to bottom
	USES_CONVERSION;
	long r;
    for (r = m_spGrid->Row + 1; r < m_spGrid->Rows; r++) {
        if (m_spGrid->GetRowHidden(r)) continue;
		LPTSTR sz = W2T(m_spGrid->GetTextMatrix(r, 2));
		if (sz && toupper(*sz) == *KeyCode) {
			m_spGrid->Select(r, m_spGrid->Cols - 1);
            m_spGrid->ShowCell(r, m_spGrid->Cols - 1);
            *KeyCode = 0;
			return;
		}
	}
    
    // not found, so look from top down to current - 1
    for (r = m_spGrid->FixedRows; r < m_spGrid->Row; r++) {
        if (m_spGrid->GetRowHidden(r)) continue;
		LPTSTR sz = W2T(m_spGrid->GetTextMatrix(r, 2));
		if (sz && toupper(*sz) == *KeyCode) {
			m_spGrid->Select(r, m_spGrid->Cols - 1);
            m_spGrid->ShowCell(r, m_spGrid->Cols - 1);
            *KeyCode = 0;
			return;
		}
	}
}

// BeforeUserResize: don't resize outline column
void _stdcall CVSPPG::OnBeforeUserResize(long Row, long Col, VARIANT_BOOL* Cancel)
{
    if (Col == 0) *Cancel = true;
}

// CellButtonClick: display color pick dialog
void _stdcall CVSPPG::OnCellButtonClick(long Row, long Col)
{
	// get property pointer
	_variant_t vProp = m_spGrid->GetRowData(Row);
	if (vProp.vt != VT_I4) return;
	CProperty* pProp = (CProperty*)V_I4(&vProp);

	// translate system colors
	COLORREF clr;
	OleTranslateColor(V_I4(&pProp->m_value), NULL, &clr);

	// show color dialog
	CHOOSECOLOR cc;
	static COLORREF clrCustom[16];
	memset(&cc, 0, sizeof(cc));
	cc.lStructSize  = sizeof(cc);
	cc.hwndOwner    = m_hWnd;
	cc.lpCustColors = (LPDWORD)clrCustom;
	cc.rgbResult    = clr;
	cc.Flags        = CC_RGBINIT;
	if (!ChooseColor(&cc)) return;
	if (cc.rgbResult == clr) return;

	// set new value
	TCHAR szColor[25];
	wsprintf(szColor, "     %#08lx", cc.rgbResult);
	m_spGrid->PutTextMatrix(Row, 3, _bstr_t(szColor));

	// mark property as dirty
	pProp->SetDirty(true);
	m_spGrid->PutCell(flexcpFontBold, Row, 3L, vtMissing, vtMissing, (long)true);

	// mark page as dirty
	SetDirty(true);
}

// DrawCell: use to show colors
void _stdcall CVSPPG::OnDrawCell(long hDC, long Row, long Col, long Left, long Top, long Right, long Bottom, VARIANT_BOOL* Done)
{
    // only need to custom draw color selection cells
    if (Col != m_spGrid->Cols - 1) return;
	_variant_t vProp = m_spGrid->GetRowData(Row);
	if (vProp.vt != VT_I4) return;
	CProperty* pProp = (CProperty*)V_I4(&vProp);
	if (pProp->m_vtNative != VT_UI4) return;

    // build rectangle for color box
    RECT rc;
	SetRect(&rc, Left + 2, Top + 2, Left + 15, Bottom - 3);
	if (rc.bottom - rc.top < 3 || rc.right - rc.left < 3) return;
    
    // translate color
	USES_CONVERSION;
	long clr = pProp->HexToLong(W2A(m_spGrid->GetTextMatrix(Row, Col)));
	OleTranslateColor(clr, NULL, (COLORREF*)&clr);
    
    // paint rectangle
	HDC hdc = (HDC)hDC;
    clr = SetBkColor(hdc, clr);
    ExtTextOut(hdc, 0, 0, ETO_OPAQUE, &rc, 0, 0, 0);
    SetBkColor(hdc, clr);
    
    // frame rectangle
    FrameRect(hdc, &rc, (HBRUSH)GetStockObject(BLACK_BRUSH));
}

