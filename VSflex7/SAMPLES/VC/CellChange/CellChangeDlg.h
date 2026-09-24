// CellChangeDlg.h : Declaration of the CCellChangeDlg

#ifndef __CELLCHANGEDLG_H_
#define __CELLCHANGEDLG_H_

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CCellChangeDlg
class CCellChangeDlg : 
	public CAxDialogImpl<CCellChangeDlg>,
	public IDispEventImpl<IDC_VSFLEXGRID1, CCellChangeDlg>
{
public:
	CCellChangeDlg()
	{
	}

	~CCellChangeDlg()
	{
	}

	enum { IDD = IDD_CELLCHANGEDLG };

BEGIN_MSG_MAP(CCellChangeDlg)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	MESSAGE_HANDLER(WM_CLOSE, OnClose)
	MESSAGE_HANDLER(WM_SIZE, OnSize)
	COMMAND_HANDLER(IDC_CHECK1, BN_CLICKED, OnCheckClicked)
	MESSAGE_HANDLER(WM_TIMER, OnTimer)
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
        // get pointer to grid
        CAxWindow wndGrid = GetDlgItem(IDC_VSFLEXGRID1);
        wndGrid.QueryControl(&m_spGrid);

        // make the grid UI-active and hook up events
        wndGrid.SetFocus();
		AtlAdviseSinkMap(this, true);

        // initialize grid
        m_spGrid->Cols = 10;
        m_spGrid->Rows = 30;
        m_spGrid->PutColFormat(-1, _bstr_t("#,###.##"));
        m_spGrid->PutTextMatrix(1, 0, _bstr_t("Count"));
        m_spGrid->PutTextMatrix(2, 0, _bstr_t("Sum"));
        m_spGrid->PutTextMatrix(3, 0, _bstr_t("Average"));
        m_spGrid->PutTextMatrix(4, 0, _bstr_t("Min"));
        m_spGrid->PutTextMatrix(5, 0, _bstr_t("Max"));
        m_spGrid->PutTextMatrix(6, 0, _bstr_t("StdDev"));
        m_spGrid->PutTextMatrix(7, 0, _bstr_t("Variance"));
        m_spGrid->FrozenRows = 7;
        m_spGrid->BackColorFrozen = 0xC0FFC0;
        m_spGrid->Editable = flexEDKbdMouse;
		m_spGrid->Redraw = flexRDBuffered;
        m_spGrid->AllowUserResizing = flexResizeColumns;
        m_spGrid->Select(m_spGrid->FixedRows + m_spGrid->FrozenRows, 1);

		// fix layout
		OnSize(uMsg, wParam, lParam, bHandled);

		return 1;  // Let the system set the focus
	}

	// adjust form layout after resizing
	LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		RECT rc;
		GetClientRect(&rc);
		InflateRect(&rc, -5, -5);
		rc.top = 25;
		CWindow wnd = GetDlgItem(IDC_VSFLEXGRID1);
		wnd.MoveWindow(&rc);
		return 0;
	}

	// done, close dialog
	LRESULT OnClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		EndDialog(IDOK);
		return 0;
	}

	// start/stop making random changes to the grid
	LRESULT OnCheckClicked(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
		BOOL bChecked = SendDlgItemMessage(IDC_CHECK1, BM_GETCHECK);
		if (bChecked)
			SetTimer(IDC_VSFLEXGRID1, 250);
		else
			KillTimer(IDC_VSFLEXGRID1);
		return 0;
	}

	// create random entries o nthe grid
	LRESULT OnTimer(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		ATLASSERT(wParam == IDC_VSFLEXGRID1);

		// randomize row and column
		long r = rand() % m_spGrid->Rows;
		if (r > m_spGrid->Rows - 1) r = m_spGrid->Rows - 1;
		if (r < m_spGrid->FixedRows + m_spGrid->FrozenRows) r = m_spGrid->FixedRows + m_spGrid->FrozenRows;
		long c = rand() % m_spGrid->Cols;
		if (c > m_spGrid->Cols - 1)  c = m_spGrid->Cols - 1;
		if (c < m_spGrid->FixedCols) c = m_spGrid->FixedCols;
    
		// assign random value (will update totals automatically)
		m_spGrid->PutTextMatrix(r, c, _bstr_t((double)(rand() % 200 - 100)));
		return 0;
	}

	// can't edit frozen rows
	VOID __stdcall OnBeforeEdit(LONG Row, LONG Col, VARIANT_BOOL * Cancel)
	{
		if (Row <= m_spGrid->FrozenRows)
			*Cancel = true;
	}

    // after a change, update summary and format changed cell
	VOID __stdcall OnCellChanged(LONG Row, LONG Col)
	{
		// not in summary area
		if (Col > 0 && Row > 6)
		{
			// format cell:
			// if the value is negative, make cell red and bold
			// if > 80, make cell green and bold
			// otherwise, reset cell format
			double v = m_spGrid->GetValueMatrix(Row, Col);
			_variant_t vRow(Row), vCol(Col);
			if (v < 0) {
				m_spGrid->PutCell(flexcpForeColor, Row, Col, Row, Col, (long)RGB(150, 0, 0));
				m_spGrid->PutCell(flexcpFontBold,  Row, Col, Row, Col, (long)true);
			} else if (v > 80) {
				m_spGrid->PutCell(flexcpForeColor, Row, Col, Row, Col, (long)RGB(0, 150, 0));
				m_spGrid->PutCell(flexcpFontBold,  Row, Col, Row, Col, (long)true);
			} else {
				m_spGrid->PutCell(flexcpCustomFormat, Row, Col, Row, Col, (long)false);
			}
        
			// update summary information
			long r1 = m_spGrid->FixedRows + m_spGrid->FrozenRows;
			long r2 = m_spGrid->Rows - 1;
			m_spGrid->PutTextMatrix(1, Col, _bstr_t(m_spGrid->GetAggregate(flexSTCount,   r1, Col, r2, Col)));
			m_spGrid->PutTextMatrix(2, Col, _bstr_t(m_spGrid->GetAggregate(flexSTSum,     r1, Col, r2, Col)));
			m_spGrid->PutTextMatrix(3, Col, _bstr_t(m_spGrid->GetAggregate(flexSTAverage, r1, Col, r2, Col)));
			m_spGrid->PutTextMatrix(4, Col, _bstr_t(m_spGrid->GetAggregate(flexSTMin,     r1, Col, r2, Col)));
			m_spGrid->PutTextMatrix(5, Col, _bstr_t(m_spGrid->GetAggregate(flexSTMax,     r1, Col, r2, Col)));
			m_spGrid->PutTextMatrix(6, Col, _bstr_t(m_spGrid->GetAggregate(flexSTStd,     r1, Col, r2, Col)));
			m_spGrid->PutTextMatrix(7, Col, _bstr_t(m_spGrid->GetAggregate(flexSTVar,     r1, Col, r2, Col)));
		}
	}


public :

BEGIN_SINK_MAP(CCellChangeDlg)
	//Make sure the Event Handlers have __stdcall calling convention
	SINK_ENTRY(IDC_VSFLEXGRID1, 0x2b, OnCellChanged)
	SINK_ENTRY(IDC_VSFLEXGRID1, 0x15, OnBeforeEdit)
END_SINK_MAP()

    // store pointer to grid control
    IVSFlexGridPtr  m_spGrid;
};

#endif //__CELLCHANGEDLG_H_
