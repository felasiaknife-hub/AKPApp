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
	COMMAND_HANDLER(IDC_BYROW, BN_CLICKED, OnByRowClicked)
	MESSAGE_HANDLER(WM_CLOSE, OnClose)
	MESSAGE_HANDLER(WM_SIZE, OnSize)
END_MSG_MAP()

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		// get pointer to control, activate it, and hook up events
		m_wndGrid = GetDlgItem(IDC_VSFLEXGRID1);
		m_wndGrid.QueryControl(&m_spGrid);
		m_wndGrid.SetFocus();
		AtlAdviseSinkMap(this, true);

		// initialize control
        m_spGrid->Editable = flexEDKbdMouse;
        m_spGrid->AllowUserResizing = flexResizeBoth;
        m_spGrid->AllowUserFreezing = flexFreezeBoth;
        m_spGrid->BackColorFrozen = RGB(200, 200, 255);
        m_spGrid->ExplorerBar = (ExplorerBarSettings)(flexExMoveRows | flexExSortShowAndMove);
        for (long r = 0; r < m_spGrid->Rows; r++) {
            for (long c = 0; c < m_spGrid->Cols; c++) {
				TCHAR sz[30];
				wsprintf(sz, "r%dc%d", r, c);
                m_spGrid->PutTextMatrix(r, c, _bstr_t(sz));
			}
		}

		// done
		return FALSE;
	}

	LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		RECT rc;
		GetClientRect(&rc);
		InflateRect(&rc, -3, -3);
		rc.top = 25;
		m_wndGrid.MoveWindow(&rc);
		return 0;
	}

	LRESULT OnClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		EndDialog(IDOK);
		return 0;
	}

	LRESULT OnByRowClicked(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
		m_spGrid->SelectionMode = (SendDlgItemMessage(IDC_BYROW, BM_GETCHECK))
			? flexSelectionListBox
			: flexSelectionFree;
		return 0;
	}


/////////////////////////////////////////////////////////////////////////////
//
// Event Handlers
//
/////////////////////////////////////////////////////////////////////////////
BEGIN_SINK_MAP(CMainDlg)
	//Make sure the Event Handlers have __stdcall calling convention
	SINK_ENTRY(IDC_VSFLEXGRID1, 0x2c, OnBeforeMoveRow)
	SINK_ENTRY(IDC_VSFLEXGRID1, 0x5, OnBeforeMouseDown)
END_SINK_MAP()

	VOID __stdcall OnBeforeMoveRow(LONG Row, LONG * Position)
	{
	    // when selection is normal, allow default behavior
		if (m_spGrid->SelectionMode != flexSelectionListBox) return;
    
		// when selection is by row, drag all selected rows to new position
        // move all selected rows to bottom: this makes it easier to
        // keep track of them as we start moving rows around
        long iSelCnt = m_spGrid->SelectedRows;
        for (long i = 0; i < iSelCnt; i++) {
			long r = m_spGrid->GetSelectedRow(0);
            m_spGrid->PutRowPosition(r, m_spGrid->Rows - 1);
			if (r < *Position) (*Position)--;
        }
        
        // move selected rows to given position
        if (*Position > Row) (*Position)++;
        for (i = 0; i < iSelCnt; i++) {
            m_spGrid->PutRowPosition(m_spGrid->Rows - 1, *Position);
		}
        
        // cancel default behavior (we have handled the move)
        *Position = Row;
	}

	VOID __stdcall OnBeforeMouseDown(SHORT Button, SHORT Shift, FLOAT X, FLOAT Y, VARIANT_BOOL * Cancel)
	{
		// allow dragging with the right mouse button on any column
		long r = m_spGrid->MouseRow;
		if (Button == 2 && r >= m_spGrid->FixedRows) {
			m_spGrid->PutCell(flexcpBackColor, r, m_spGrid->FixedCols, r, m_spGrid->Cols - 1L, 255L);
			r = m_spGrid->DragRow(r);
			m_spGrid->PutCell(flexcpCustomFormat, r, m_spGrid->FixedCols, r, m_spGrid->Cols - 1L, 0L);
		}
	}


/////////////////////////////////////////////////////////////////////////////
//
// Members
//
/////////////////////////////////////////////////////////////////////////////
	CAxWindow		m_wndGrid;
	IVSFlexGridPtr	m_spGrid;
};

#endif //__CMAINDLG_H_
