// CustDataDlg.h : Declaration of the CCustDataDlg

#ifndef __CUSTDATADLG_H_
#define __CUSTDATADLG_H_

#include "resource.h"       // main symbols


/////////////////////////////////////////////////////////////////////////////
// CCustDataDlg
class CCustDataDlg : 
	public CAxDialogImpl<CCustDataDlg>,
	public IDispEventImpl<IDC_FGLINKS, CCustDataDlg>,
	public IDispEventImpl<IDC_FGNODES, CCustDataDlg>
{
public:
	CCustDataDlg()
	{
	}

	~CCustDataDlg()
	{
	}

	enum { IDD = IDD_CUSTDATADLG };

BEGIN_MSG_MAP(CCustDataDlg)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	MESSAGE_HANDLER(WM_CLOSE, OnClose)
	MESSAGE_HANDLER(WM_SIZE, OnSize)
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

BEGIN_SINK_MAP(CCustDataDlg)
	//Make sure the Event Handlers have __stdcall calling convention
	SINK_ENTRY(IDC_FGLINKS, 0x15, OnBeforeEditLinks)
	SINK_ENTRY(IDC_FGLINKS, 0x18, OnAfterEditLinks)
	SINK_ENTRY(IDC_FGNODES, 0x18, OnAfterEditNodes)
END_SINK_MAP()

////////////////////////////////////////////////////////////////////////////////
//
// Message Handlers
//
////////////////////////////////////////////////////////////////////////////////

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		// initialize label controls
		m_fntLabel.Create(_T("Tahoma"), 10, true);
		m_wndLblNodes = GetDlgItem(IDC_LBLNODES);
		m_wndLblLinks = GetDlgItem(IDC_LBLLINKS);
		m_wndLblNodes.SetFont(m_fntLabel.GetFontHandle());
		m_wndLblLinks.SetFont(m_fntLabel.GetFontHandle());

		// get pointers to grids
		m_wndLinks = GetDlgItem(IDC_FGLINKS);
		m_wndLinks.QueryControl(&m_spLinks);
		m_wndLinks.SetFocus();
		m_wndNodes = GetDlgItem(IDC_FGNODES);
		m_wndNodes.QueryControl(&m_spNodes);
		m_wndNodes.SetFocus();

		// start getting events
		AtlAdviseSinkMap(this, true);

		// combo list for link and node names
		_bstr_t strNameList("|Leg|Boom|Girder|Trolley|Spreader|Cabin|Utility|General");

		// initialize custom data sources
		m_FlexDSNodes.m_pData = &m_TheData;
		m_FlexDSLinks.m_pData = &m_TheData;

		// assign custom data sources to control
		IVSFlexDataSourcePtr spDSNodes((IUnknown*)&m_FlexDSNodes);
        m_spNodes->FlexDataSource = spDSNodes;

		IVSFlexDataSourcePtr spDSLinks((IUnknown*)&m_FlexDSLinks);
        m_spLinks->FlexDataSource = spDSLinks;
		// initialize node grid
		m_spNodes->Editable = flexEDKbdMouse;
		m_spNodes->Redraw = flexRDBuffered;
        m_spNodes->FrozenCols = 1;
        m_spNodes->PutColWidth(0, 700);
        m_spNodes->PutColWidth(1, 1200);
        m_spNodes->AllowUserResizing = flexResizeBothUniform;
        m_spNodes->AllowUserFreezing = flexFreezeBoth;
        m_spNodes->BackColorFrozen = 0xFFC0C0;
		_bstr_t fmt(_T("#,###.00"));
        for (int i = 2; i <= 4; i++)
			m_spNodes->PutColFormat(i, fmt);
        for (i = 2; i < m_spNodes->Cols; i++)
            m_spNodes->PutColAlignment(i, flexAlignRightTop);
		m_spNodes->PutColComboList(1, strNameList);

		// initialize link grid
		m_spLinks->Editable = flexEDKbdMouse;
		m_spLinks->Redraw = flexRDBuffered;
        m_spLinks->FrozenCols = 1;
        m_spLinks->PutColWidth(0, 700);
        m_spLinks->PutColWidth(1, 1200);
        m_spLinks->AllowUserResizing = flexResizeBothUniform;
        m_spLinks->AllowUserFreezing = flexFreezeBoth;
        m_spLinks->BackColorFrozen = 0xFFC0C0;
        for (i = 2; i < m_spLinks->Cols; i++)
            m_spLinks->PutColAlignment(i, flexAlignRightTop);
		for (i = 4; i <= 8; i++)
			m_spLinks->PutColFormat(i, fmt);
		m_spLinks->PutColFormat(5, _bstr_t("#,###"));
		m_spLinks->PutColComboList(1, strNameList);

		return 0;
	}

	LRESULT OnClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		EndDialog(IDOK);
		return 0;
	}

	LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		// get client rect
		RECT rc, rcClient;
		GetClientRect(&rcClient);
		InflateRect(&rcClient, -4, -4);
		int iCliHei = rcClient.bottom - rcClient.top;

		// get label height
		m_wndLblNodes.GetClientRect(&rc);
		int iLblHei = rc.bottom - rc.top;

		// move nodes label and grid
		rc = rcClient;
		rc.bottom = rc.top + iLblHei;
		m_wndLblNodes.MoveWindow(&rc);
		rc.top = rc.bottom;
		rc.bottom = rc.top + iCliHei/2 - iLblHei;
		m_wndNodes.MoveWindow(&rc);

		// move links label and grid
		rc.top = rc.bottom;
		rc.bottom = rc.top + iLblHei;
		m_wndLblLinks.MoveWindow(&rc);
		rc.top = rc.bottom;
		rc.bottom = rcClient.bottom;
		m_wndLinks.MoveWindow(&rc);

		// done
		return 0;
	}

////////////////////////////////////////////////////////////////////////////////
//
// Event Handlers
//
////////////////////////////////////////////////////////////////////////////////

	VOID __stdcall OnBeforeEditLinks(LONG Row, LONG Col, VARIANT_BOOL * Cancel)
	{
	    // cannot edit link length
		if (Col == 4)
			*Cancel = true;
	}
	VOID __stdcall OnAfterEditLinks(LONG Row, LONG Col)
	{
	    // after changing link connectivity, show new length by refreshing a single cell
	    if (Col == 2 || Col == 3)
			m_spLinks->PutCell(flexcpRefresh, Row, 4L, vtMissing, vtMissing, 1L);
	}

	VOID __stdcall OnAfterEditNodes(LONG Row, LONG Col)
	{
	    // after moving nodes, update link lengths
		// note that setting the Redraw property automatically refreshes the control
		if (Col > 1)
			m_spLinks->Redraw = flexRDBuffered;
	}


////////////////////////////////////////////////////////////////////////////////
//
// Member Variables
//
////////////////////////////////////////////////////////////////////////////////
private:

	CGDIFont		m_fntLabel;
	IVSFlexGridPtr	m_spNodes, m_spLinks;
	CAxWindow		m_wndNodes, m_wndLinks;
	CWindow			m_wndLblNodes, m_wndLblLinks;
	
	CTheData		m_TheData;
	CFlexDSLinks	m_FlexDSLinks;
	CFlexDSNodes	m_FlexDSNodes;
};

#endif //__CUSTDATADLG_H_
