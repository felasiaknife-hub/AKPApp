// VSPPGFlex.h : Declaration of the CVSPPGFlex

#ifndef __VSPPGFLEX_H_
#define __VSPPGFLEX_H_

#include "resource.h"       // main symbols

EXTERN_C const CLSID CLSID_VSPPGFlex;

/////////////////////////////////////////////////////////////////////////////
// CVSPPGFlex global definitions

#define MSG_NOFLEX			_T("** Cannot create property grid.\n** Please make sure that the VSFLEX7L.OCX library is registered on this computer.")
#define STYLE_HOST			(WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS)
#define IDC_GRID			666
#define IDC_PREVIEW			667
#define BORDER				3

/////////////////////////////////////////////////////////////////////////////
// CVSPPGFlex
class ATL_NO_VTABLE CVSPPGFlex :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVSPPGFlex, &CLSID_VSPPGFlex>,
	public IPropertyPageImpl<CVSPPGFlex>,
	public CDialogImpl<CVSPPGFlex>,
	public IDispEventImpl<IDC_GRID, CVSPPGFlex>,	// needed to sink events
	public IDispEventImpl<IDC_PREVIEW, CVSPPGFlex>	// needed to sink events
{
public:
	CVSPPGFlex() 
	{
		m_dwTitleID = IDS_TITLEVSPPGFlex;
		m_dwHelpFileID = IDS_HELPFILE;
		m_dwDocStringID = IDS_DOCSTRINGVSPPGFlex;
	}

	enum {IDD = IDD_VSPPGFLEX};

DECLARE_REGISTRY_RESOURCEID(IDR_VSPPGFLEX)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CVSPPGFlex) 
	COM_INTERFACE_ENTRY(IPropertyPage)
END_COM_MAP()

BEGIN_MSG_MAP(CVSPPGFlex)
	CHAIN_MSG_MAP(IPropertyPageImpl<CVSPPGFlex>)
	MESSAGE_HANDLER(WM_INITDIALOG,	OnInitDialog)
	MESSAGE_HANDLER(WM_DESTROY,		OnDestroy)
	MESSAGE_HANDLER(WM_MOUSEMOVE,	OnMouseMove)
	MESSAGE_HANDLER(WM_LBUTTONDOWN,	OnMouseDown)
	MESSAGE_HANDLER(WM_SETCURSOR,   OnSetCursor)
	COMMAND_HANDLER(IDC_CLEAR,  BN_CLICKED, OnClickedClear)
	COMMAND_HANDLER(IDC_COLINS, BN_CLICKED, OnClickedInsert)
	COMMAND_HANDLER(IDC_COLDEL, BN_CLICKED, OnClickedDelete)
END_MSG_MAP()

// event map (see event dispids in .tli file)
BEGIN_SINK_MAP(CVSPPGFlex)

	// events handled by the property editor
	SINK_ENTRY(IDC_GRID,	DISPID_DBLCLICK,	OnDblClick)				// cycle list items
	SINK_ENTRY(IDC_GRID,	DISPID_KEYDOWN,		OnKeyDown)				// search for property
	SINK_ENTRY(IDC_GRID,	6,					OnBeforeRowColChange)	// only the value column can be selected
	SINK_ENTRY(IDC_GRID,	21,					OnBeforeEdit)			// set ComboList
	SINK_ENTRY(IDC_GRID,	22,					OnStartEdit)			// cycle list items
	SINK_ENTRY(IDC_GRID,	24,					OnAfterEdit)			// apply new settings, mark as dirty

	// events handled by the preview window
	SINK_ENTRY(IDC_PREVIEW,	DISPID_DBLCLICK,	OnDblClickPV)			// dbl-clicks start editing fixed cells
	SINK_ENTRY(IDC_PREVIEW,	1,					OnSelChangePV)			// update help and property window values
	SINK_ENTRY(IDC_PREVIEW,	24,					OnAfterEditPV)			// autosize fixed row height
	SINK_ENTRY(IDC_PREVIEW,	16,					OnAfterMoveColumnPV)	// moved column, we're dirty
	SINK_ENTRY(IDC_PREVIEW,	18,					OnAfterUserResizePV)	// show new width value

END_SINK_MAP()

// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	// interface methods
	STDMETHOD(Apply)();
	STDMETHOD(SetObjects)(ULONG nObjects, IUnknown **ppUnk);
	STDMETHOD(TranslateAccelerator)(MSG *pMsg);

	// message handlers
	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnDestroy   (UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnMouseMove (UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnMouseDown (UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnSetCursor (UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnClickedClear (WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	LRESULT OnClickedInsert(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	LRESULT OnClickedDelete(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);

	// event handlers for property control (must be declared with __stdcall)
	void _stdcall OnDblClick();
	void _stdcall OnKeyDown(short* KeyCode, short Shift);
	void _stdcall OnBeforeRowColChange(long OldRow, long OldCol, long NewRow, long NewCol, VARIANT_BOOL* Cancel);
	void _stdcall OnRowColChange();
	void _stdcall OnBeforeEdit(long Row, long Col, VARIANT_BOOL* Cancel);
	void _stdcall OnStartEdit(long Row, long Col, VARIANT_BOOL* Cancel);
	void _stdcall OnAfterEdit(long Row, long Col);

	// event handlers for preview control (must be declared with __stdcall)
	void _stdcall OnDblClickPV();
	void _stdcall OnSelChangePV();
	void _stdcall OnAfterEditPV(long Row, long Col);
	void _stdcall OnAfterMoveColumnPV(long Col, long *Position);
	void _stdcall OnAfterUserResizePV(long Row, long Col);

	// internal methods
	void	InitGrid();
	void	InitGridPreview();
	int		GetFontHeight();
	void	GetGridRect(LPRECT prc, BOOL bGrid = true);
	int		CountLinesInCell(long Row, long Col);
	int		CountCellsInRow(long Row);
    int     CountRowsInGrid();
	LPTSTR	InStr(LPTSTR str1, LPTSTR str2);
	bool	SelectNextItem(long Row, long Col);

	// members
	CComPtr<IFont>	m_spFont;			// font to be used in the controls
	CAxWindow		m_wndHost;			// ActiveX control host
	IVSFlexGridPtr	m_spGrid;			// smart pointer to FlexGrid control
	CAxWindow		m_wndHostPreview;	// ActiveX control host
	IVSFlexGridPtr	m_spGridPreview;	// smart pointer to FlexGrid control
	CWindow			m_wndHelp;			// property description window
	BOOL			m_bDblClick;		// editing started with double-click?
	CComDispatchDriver m_spDisp;		// dispatch driver for control being edited (to handle all flavors)
	HICON			m_hIconClr;  		// icon for clear button
	HICON			m_hIconIns;  		// icon for insert column button
	HICON			m_hIconDel;  		// icon for delete column button

	// dispatch helpers
	HRESULT GetCtlProp(LPCTSTR szName, VARIANT* pVal)
	{
		DISPID dwDispID;
		HRESULT hr = m_spDisp.GetIDOfName(_bstr_t(szName), &dwDispID);
		if (FAILED(hr)) {
			ATLASSERT(false);
			return E_INVALIDARG;
		}
		DISPPARAMS dParams = {NULL, NULL, 0, 0};
		return m_spDisp->Invoke(dwDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
								&dParams, pVal, NULL, NULL);
	}

	HRESULT GetCtlProp(LPCTSTR szName, long iCol, VARIANT* pVal)
	{
		DISPID dwDispID;
		HRESULT hr = m_spDisp.GetIDOfName(_bstr_t(szName), &dwDispID);
		if (FAILED(hr)) {
			ATLASSERT(false);
			return E_INVALIDARG;
		}
		_variant_t vCol(iCol);
		DISPPARAMS dParams = {NULL, NULL, 0, 0};
		dParams.rgvarg = &vCol;
		dParams.cArgs = 1;
		return m_spDisp->Invoke(dwDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
								&dParams, pVal, NULL, NULL);
	}

	HRESULT GetCtlProp(LPCTSTR szName, long iRow, long iCol, VARIANT* pVal)
	{
		DISPID dwDispID;
		HRESULT hr = m_spDisp.GetIDOfName(_bstr_t(szName), &dwDispID);
		if (FAILED(hr)) {
			ATLASSERT(false);
			return E_INVALIDARG;
		}
		_variant_t vParms[2];
		vParms[0] = iCol;
		vParms[1] = iRow;
		DISPPARAMS dParams = {NULL, NULL, 0, 0};
		dParams.rgvarg = vParms;
		dParams.cArgs = 2;
		return m_spDisp->Invoke(dwDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
								&dParams, pVal, NULL, NULL);
	}

	long GetCtlPropLong(LPCTSTR szName)
	{
		_variant_t var = 0L;
		GetCtlProp(szName, &var);
		var.ChangeType(VT_I4);
		return V_I4(&var);
	}

	long GetCtlPropLong(LPCTSTR szName, long iCol)
	{
		_variant_t var = 0L;
		GetCtlProp(szName, iCol, &var);
		var.ChangeType(VT_I4);
		return V_I4(&var);
	}

	long GetCtlPropLong(LPCTSTR szName, long iRow, long iCol)
	{
		_variant_t var = 0L;
		GetCtlProp(szName, iRow, iCol, &var);
		var.ChangeType(VT_I4);
		return V_I4(&var);
	}

	_bstr_t GetCtlPropStr(LPCTSTR szName)
	{
		_variant_t var = "";
		GetCtlProp(szName, &var);
		var.ChangeType(VT_BSTR);
	    return _bstr_t(var);
	}

	_bstr_t GetCtlPropStr(LPCTSTR szName, long iCol)
	{
		_variant_t var = "";
		GetCtlProp(szName, iCol, &var);
		var.ChangeType(VT_BSTR);
	    return _bstr_t(var);
	}

	_bstr_t GetCtlPropStr(LPCTSTR szName, long iRow, long iCol)
	{
		_variant_t var = "";
		GetCtlProp(szName, iRow, iCol, &var);
		var.ChangeType(VT_BSTR);
	    return _bstr_t(var);
	}

	IFontDisp* GetCtlPropFont()
	{
		_variant_t vVal;
		m_spDisp.GetProperty(DISPID_FONT, &vVal);
		CComQIPtr<IFontDisp, &IID_IFontDisp> spFont(V_DISPATCH(&vVal));
		return spFont;
	}

    HRESULT SetCtlProp(LPCTSTR szName, _variant_t& var)
	{
        return m_spDisp.PutPropertyByName(_bstr_t(szName), &var);
	}
};

#endif //__VSPPGFLEX_H_
