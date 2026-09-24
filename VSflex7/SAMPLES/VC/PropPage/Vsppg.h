// VSPPG.h : Declaration of the CVSPPG

#ifndef __VSPPG_H_
#define __VSPPG_H_

#include "CProperty.h"		// control property class
#include "resource.h"       // main symbols

EXTERN_C const CLSID CLSID_VSPPG;


/////////////////////////////////////////////////////////////////////////////
// CVSPPG global definitions

extern  HICON				g_hIcon;
#define MSG_NOFLEX			_T("** Cannot create property grid.\n** Please make sure that the VSFLEX7L.OCX library is registered on this computer.")
#define STYLE_HOST			(WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS)
#define IDC_GRID			666
#define BORDER				3


/////////////////////////////////////////////////////////////////////////////
// CVSPPG
class ATL_NO_VTABLE CVSPPG :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CVSPPG, &CLSID_VSPPG>,
	public IPropertyPageImpl<CVSPPG>,
	public CDialogImpl<CVSPPG>,
	public IDispEventImpl<IDC_GRID, CVSPPG>	// needed to sink events
{
public:
	CVSPPG() 
	{
		m_dwTitleID = IDS_TITLEVSPPG;
		m_dwHelpFileID = IDS_HELPFILE;
		m_dwDocStringID = IDS_DOCSTRINGVSPPG;
	}

	enum {IDD = IDD_VSPPG};

DECLARE_REGISTRY_RESOURCEID(IDR_VSPPG)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CVSPPG) 
	COM_INTERFACE_ENTRY(IPropertyPage)
END_COM_MAP()

BEGIN_MSG_MAP(CVSPPG)
	CHAIN_MSG_MAP(IPropertyPageImpl<CVSPPG>)
	MESSAGE_HANDLER(WM_INITDIALOG,	OnInitDialog)
	MESSAGE_HANDLER(WM_DESTROY,		OnDestroy)
	COMMAND_HANDLER(IDC_BYCAT, BN_CLICKED, OnClickedByCat)
END_MSG_MAP()

// event map (see event dispids in .tli file)
BEGIN_SINK_MAP(CVSPPG)
	SINK_ENTRY(IDC_GRID,	DISPID_DBLCLICK,	OnDblClick)				// collapse/expand outline, cycle list items
	SINK_ENTRY(IDC_GRID,	DISPID_KEYDOWN,		OnKeyDown)				// collapse/expand, navigate
	SINK_ENTRY(IDC_GRID,	6,					OnBeforeRowColChange)	// restrict to values
	SINK_ENTRY(IDC_GRID,	2,					OnRowColChange)			// update help text
	SINK_ENTRY(IDC_GRID,	21,					OnBeforeEdit)			// set ComboList
	SINK_ENTRY(IDC_GRID,	22,					OnStartEdit)			// change selection on dbl-click
	SINK_ENTRY(IDC_GRID,	24,					OnAfterEdit)			// mark as dirty
	SINK_ENTRY(IDC_GRID,	17,					OnBeforeUserResize)		// can't resize outline col
	SINK_ENTRY(IDC_GRID,	34,					OnCellButtonClick)		// show color dialog
	SINK_ENTRY(IDC_GRID,	32,					OnDrawCell)				// show color property
END_SINK_MAP()

// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	// interface methods
	STDMETHOD(Apply)();
	STDMETHOD(TranslateAccelerator)(MSG *pMsg);

	// message handlers
	LRESULT OnInitDialog  (UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnDestroy     (UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnClickedByCat(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);

	// internal methods
	void	InitGrid();
	int		GetFontHeight();
	HRESULT GetEnumTypeInfo(ITypeInfo *pTI, HREFTYPE hrt, ITypeInfo** ppEnumInfo);
	void	FillPropertyList();
	void	ClearPropertyList();
	void    DisplayCategorized();
	void    DisplayAlphabetic();
	LPTSTR	InStr(LPTSTR str1, LPTSTR str2);
	bool	SelectNextItem(long Row, long Col);

	// event handlers (must be declared with __stdcall)
	void _stdcall OnDblClick();
	void _stdcall OnKeyDown(short* KeyCode, short Shift);
	void _stdcall OnBeforeRowColChange(long OldRow, long OldCol, long NewRow, long NewCol, VARIANT_BOOL* Cancel);
	void _stdcall OnRowColChange();
	void _stdcall OnAfterEdit(long Row, long Col);
	void _stdcall OnBeforeEdit(long Row, long Col, VARIANT_BOOL* Cancel);
	void _stdcall OnStartEdit(long Row, long Col, VARIANT_BOOL * Cancel);

	void _stdcall OnBeforeUserResize(long Row, long Col, VARIANT_BOOL* Cancel);
	void _stdcall OnCellButtonClick(long Row, long Col);
	void _stdcall OnDrawCell(long hDC, long Row, long Col, long Left, long Top, long Right, long Bottom, VARIANT_BOOL* Done);

	// members
	CComPtr<IFont>	m_spFont;			// font to be used in the controls
	CAxWindow		m_wndHost;			// ActiveX control host
	IVSFlexGridPtr	m_spGrid;			// smart pointer to FlexGrid control
	CWindow			m_wndHelp;			// property description window
	BOOL			m_bHasCat;			// control can categorize properties?
	HICON			m_hIcon;			// icon for button control
	BOOL			m_bDblClick;		// editing started with double-click?
};

#endif //__VSPPG_H_
