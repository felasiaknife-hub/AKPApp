// SelectionDlg.h : Declaration of the CSelectionDlg

#ifndef __SELECTIONDLG_H_
#define __SELECTIONDLG_H_

#include "resource.h"       // main symbols


/////////////////////////////////////////////////////////////////////////////
// CSelectionDlg
class CSelectionDlg : 
	public CAxDialogImpl<CSelectionDlg>,
	public IDispEventImpl<IDC_VSFLEXGRID, CSelectionDlg>
{
public:
	CSelectionDlg()
	{
	}

	~CSelectionDlg()
	{
	}

	enum { IDD = IDD_SELECTIONDLG };

BEGIN_MSG_MAP(CSelectionDlg)
	MESSAGE_HANDLER(WM_INITDIALOG,	OnInitDialog)
	MESSAGE_HANDLER(WM_CLOSE,		OnClose)
	MESSAGE_HANDLER(WM_SIZE,		OnSize)
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

BEGIN_SINK_MAP(CSelectionDlg)
	//Make sure the Event Handlers have __stdcall calling convention
	SINK_ENTRY(IDC_VSFLEXGRID, 0x6, OnBeforeRowColChangeGrid)
	SINK_ENTRY(IDC_VSFLEXGRID, 0x8, OnBeforeSelChangeGrid)
	SINK_ENTRY(IDC_VSFLEXGRID, 0xa, OnBeforeScrollGrid)
	SINK_ENTRY(IDC_VSFLEXGRID, 0xb, OnAfterScrollGrid)
END_SINK_MAP()

////////////////////////////////////////////////////////////////////////////////
//
// Message Handlers
//
////////////////////////////////////////////////////////////////////////////////

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
        // get pointer to grid
		m_wndGrid = GetDlgItem(IDC_VSFLEXGRID);
		m_wndGrid.QueryControl(&m_spGrid);

        // make the grid UI-active and hook up events
		m_wndGrid.SetFocus();
		AtlAdviseSinkMap(this, true);

		// initialize grid
		m_spGrid->Rows = 30;
		m_spGrid->Editable = flexEDKbdMouse;
		for (long r = m_spGrid->FixedRows; r < m_spGrid->Rows; r++) {
			for (long c = m_spGrid->FixedCols; c < m_spGrid->Cols; c++) {
				TCHAR sz[20];
				wsprintf(sz, "r%dc%d", r, c);
				m_spGrid->PutTextMatrix(r, c, _bstr_t(sz));
			}
		}
    
		// highlight protected range
		m_spGrid->PutCell(flexcpBackColor, 2L, 2L, 6L, 6L, (long)RGB(200, 250, 200));

		// done
		return 0;
	}

	LRESULT OnClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		EndDialog(IDOK);
		return 0;
	}

	LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		RECT rc;
		GetClientRect(&rc);
		InflateRect(&rc, -4, -4);
		rc.top = 55;
		m_wndGrid.MoveWindow(&rc);
		return 0;
	}


////////////////////////////////////////////////////////////////////////////////
//
// Event Handlers
//
////////////////////////////////////////////////////////////////////////////////

	VOID __stdcall OnBeforeRowColChangeGrid(LONG OldRow, LONG OldCol, LONG NewRow, LONG NewCol, VARIANT_BOOL * Cancel)
	{
		// cancel if new cell is in protected area
		if (SendDlgItemMessage(IDC_PROTECT, BM_GETCHECK)) {
			if (NewRow >= 2 && NewRow <= 6 && NewCol >= 2 && NewCol <= 6)
				*Cancel = true;
		}
	}

	VOID __stdcall OnBeforeSelChangeGrid(LONG OldRowSel, LONG OldColSel, LONG NewRowSel, LONG NewColSel, VARIANT_BOOL * Cancel)
	{
		// cancel non-vertical selection
		if (SendDlgItemMessage(IDC_VERTICAL, BM_GETCHECK)) {
			if (NewColSel != m_spGrid->ColSel)
				*Cancel = true;
		}
    
		// cancel if new selection is on protected area
		if (SendDlgItemMessage(IDC_PROTECT, BM_GETCHECK)) {
			if (m_spGrid->Row < 2 && NewRowSel < 2) return;
			if (m_spGrid->Col < 2 && NewColSel < 2) return;
			if (m_spGrid->Row > 6 && NewRowSel > 6) return;
			if (m_spGrid->Col > 6 && NewColSel > 6) return;
			*Cancel = true;
		}
	}

	VOID __stdcall OnBeforeScrollGrid(LONG OldTopRow, LONG OldLeftCol, LONG NewTopRow, LONG NewLeftCol, VARIANT_BOOL * Cancel)
	{
		// cancel if editing
		if (SendDlgItemMessage(IDC_EDITING, BM_GETCHECK)) {
			if (m_spGrid->EditWindow != 0)
				*Cancel = true;
		}
	}

	VOID __stdcall OnAfterScrollGrid(LONG OldTopRow, LONG OldLeftCol, LONG NewTopRow, LONG NewLeftCol)
	{
		// keep current cell visible
		if (SendDlgItemMessage(IDC_CURRENT, BM_GETCHECK))
			m_spGrid->ShowCell(m_spGrid->Row, m_spGrid->Col);
	}


////////////////////////////////////////////////////////////////////////////////
//
// Member Variables
//
////////////////////////////////////////////////////////////////////////////////

	CAxWindow		m_wndGrid;
	IVSFlexGridPtr	m_spGrid;
};

#endif //__SELECTIONDLG_H_
