// CMainDlg.h : Declaration of the CMainDlg

#ifndef __CMAINDLG_H_
#define __CMAINDLG_H_

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CMainDlg
class CMainDlg : 
	public CAxDialogImpl<CMainDlg>,
	public IDispEventImpl<IDC_VSFLEXGRID1, CMainDlg>
{
public:
	CMainDlg()
	{
	}

	~CMainDlg()
	{
	}

	enum { IDD = IDD_CMAINDLG };


/////////////////////////////////////////////////////////////////////////////
//
// Message Handlers
//
/////////////////////////////////////////////////////////////////////////////
BEGIN_MSG_MAP(CMainDlg)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	MESSAGE_HANDLER(WM_SIZE, OnSize)
	MESSAGE_HANDLER(WM_CLOSE, OnClose)
END_MSG_MAP()

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		// get pointer to control, activate it, and hook up events
		m_wndGrid = GetDlgItem(IDC_VSFLEXGRID1);
		m_wndGrid.QueryControl(&m_spGrid);
		m_wndGrid.SetFocus();
		AtlAdviseSinkMap(this, true);

		// initialize control
        m_spGrid->Rows = 1;
        m_spGrid->Cols = 2;
        m_spGrid->FixedCols = 0;
        m_spGrid->ExtendLastCol = true;
        m_spGrid->PutTextMatrix(0, 0, _bstr_t("Tag"));
        m_spGrid->PutTextMatrix(0, 1, _bstr_t("Value"));
        m_spGrid->PutColAlignment(-1, flexAlignLeftTop);
                
        // outline
        m_spGrid->OutlineCol = 0;
        m_spGrid->OutlineBar = flexOutlineBarSimpleLeaf;
        m_spGrid->MergeCells = flexMergeOutline;
        
        // other
        m_spGrid->AllowUserResizing = flexResizeColumns;
        m_spGrid->AllowSelection = false;
        m_spGrid->HighLight = flexHighlightNever;
        m_spGrid->GridLines = flexGridFlatVert;

		// build XML file name
		TCHAR szFileName[512];
		GetModuleFileName(_Module.GetModuleInstance(), szFileName, 512);
		LPTSTR sz = _tcsrchr(szFileName, '\\');
		if (sz) sz[1] = 0;
		_tcscat(szFileName, "test.xml");
		HFILE f = _lopen(szFileName, OF_READ);
		if (HFILE_ERROR == f) {
			sz = _tcsrchr(szFileName, '\\');
			if (sz) *sz = 0;
			sz = _tcsrchr(szFileName, '\\');
			if (sz) sz[1] = 0;
			_tcscat(szFileName, "test.xml");
			f = _lopen(szFileName, OF_READ);
		}
		if (HFILE_ERROR == f) {
			MessageBox("Can't find sample XML file (test.xml).\r\n"
				       "Please copy it to the application directory.",
					   "File not found",
					   MB_ICONERROR);
			return FALSE;
		}
	    _lclose(f);

		// read XML document into the grid
		MSXML::IXMLDOMDocumentPtr spXML;
		if (FAILED(spXML.CreateInstance(__uuidof(MSXML::DOMDocument)))) {
			MessageBox("Cannot create XML parser. Please make sure MSXML.DLL is registered.", "File not found", MB_ICONERROR);
			return FALSE;
		}
		if (FALSE == spXML->load(_bstr_t(szFileName))) {
			MessageBox("Cannot open test file 'test.xml'.", "File not found", MB_ICONERROR);
			return FALSE;
		}
		m_spGrid->Redraw = flexRDNone;
		DumpXMLNode(spXML->childNodes->item[1], 0);
		m_spGrid->AutoSize(0L, 1L, 0L, 200L);
		m_spGrid->Redraw = flexRDBuffered;

		// done
		return FALSE;
	}

	LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		RECT rc;
		GetClientRect(&rc);
		InflateRect(&rc, -3, -3);
		m_wndGrid.MoveWindow(&rc);
		return 0;
	}

	LRESULT OnClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		EndDialog(IDOK);
		return 0;
	}


/////////////////////////////////////////////////////////////////////////////
//
// Event Handlers
//
/////////////////////////////////////////////////////////////////////////////
BEGIN_SINK_MAP(CMainDlg)
	//Make sure the Event Handlers have __stdcall calling convention
	SINK_ENTRY(IDC_VSFLEXGRID1, DISPID_KEYDOWN, OnKeyDown)
END_SINK_MAP()

	VOID __stdcall OnKeyDown(SHORT * KeyCode, SHORT Shift)
	{
		if (m_spGrid->GetIsSubtotal(m_spGrid->Row)) {
			switch (*KeyCode) {
				case VK_LEFT:
					m_spGrid->PutIsCollapsed(m_spGrid->Row, flexOutlineCollapsed);
                    break;
				case VK_RIGHT:
					m_spGrid->PutIsCollapsed(m_spGrid->Row, flexOutlineExpanded);
                    break;
			}
		}
	}


/////////////////////////////////////////////////////////////////////////////
//
// Utilities
//
/////////////////////////////////////////////////////////////////////////////
	void DumpXMLNode(MSXML::IXMLDOMNodePtr spNode, int lvl)
	{
		// count children
		int iChildren = spNode->childNodes->length;
    
		// single child? add leaf node
		if (iChildren == 1) {
			_bstr_t strItem = spNode->nodeName;
			strItem += "\t";
			strItem += spNode->text;
			m_spGrid->AddItem(strItem);
			m_spGrid->PutCell(flexcpBackColor, m_spGrid->Rows - 1L, 1L, vtMissing, vtMissing, 0xf0f0f0L);
			m_spGrid->PutIsSubtotal(m_spGrid->Rows - 1, true);
			m_spGrid->PutRowOutlineLevel(m_spGrid->Rows - 1, lvl);
			return;
		}
    
		// many children? add branch node
		m_spGrid->AddItem(spNode->nodeName);
		m_spGrid->PutIsSubtotal(m_spGrid->Rows - 1, true);
		m_spGrid->PutRowOutlineLevel(m_spGrid->Rows - 1, lvl);
    
		// dump child nodes
		for (int i = 0; i < iChildren; i++) {
			DumpXMLNode(spNode->childNodes->item[i], lvl + 1);
		}
	}
    

/////////////////////////////////////////////////////////////////////////////
//
// Member Variables
//
/////////////////////////////////////////////////////////////////////////////
	CAxWindow		m_wndGrid;
	IVSFlexGridPtr	m_spGrid;
};

#endif //__CMAINDLG_H_
