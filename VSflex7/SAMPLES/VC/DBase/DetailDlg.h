// DetailDlg.h : Declaration of the CDetailDlg

#ifndef __DETAILDLG_H_
#define __DETAILDLG_H_

#include "resource.h"       // main symbols
#include <atlhost.h>

/////////////////////////////////////////////////////////////////////////////
// CDetailDlg
class CDetailDlg : 
	public CAxDialogImpl<CDetailDlg>,
	public IDispEventImpl<IDC_FGDETAIL, CDetailDlg>
{
public:
	CDetailDlg()
	{
	}

	~CDetailDlg()
	{
	}

	enum { IDD = IDD_DETAILDLG };


////////////////////////////////////////////////////////////////////////////////////////
//
// Message Handlers
//
////////////////////////////////////////////////////////////////////////////////////////
BEGIN_MSG_MAP(CDetailDlg)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	MESSAGE_HANDLER(WM_SIZE, OnSize)
	MESSAGE_HANDLER(WM_ACTIVATE, OnActivate)
	MESSAGE_HANDLER(WM_SHOWWINDOW, OnShowWindow)
END_MSG_MAP()

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		// get pointer to VSFlexGrid control
		m_wndGrid = GetDlgItem(IDC_FGDETAIL);
		m_wndGrid.QueryControl(&m_spGrid);

		// activate control and events
		m_wndGrid.SetFocus();
		AtlAdviseSinkMap(this, TRUE);

		// create imagelists we'll bind to grid columns
		m_imgDBCursor.Create(IDB_CURSOR, 16, 0, RGB(255,255,255));

		// more initialization on the AfterDataRefresh event
		m_spGrid->AllowSelection = false;
		m_spGrid->AllowUserResizing = flexResizeBoth;
		m_spGrid->AllowUserFreezing = flexFreezeBoth;
		m_spGrid->BackColorFrozen = 14737632;
		m_spGrid->GridLines = flexGridFlatVert;
		m_spGrid->ExtendLastCol = true;
		m_spGrid->ScrollTips = true;

		return 1;  // Let the system set the focus
	}

	LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		RECT rc;
		GetClientRect(&rc);
		m_wndGrid.MoveWindow(&rc);
		return 1;
	}

	LRESULT OnActivate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		// deactivating? hide yourself!
        if (LOWORD(wParam) == WA_INACTIVE) {
            ShowWindow(SW_HIDE);
            m_spGrid->DataSource = NULL;
        }
		return 0;
	}

	LRESULT OnShowWindow(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		// showing? update recordset
		BOOL fShow = (BOOL)wParam;		// show/hide flag 
		int  fnStatus = (int)lParam;	// status flag 
		if (fShow != TRUE || fnStatus != 0) return 0;

		// sanity
		ATLASSERT(m_spGridMain != NULL && m_spConnMain != NULL);

		// get current product id
		long lCol = m_spGridMain->GetColIndex(_bstr_t("Product"));
		_bstr_t strProdID = m_spGridMain->GetTextMatrix(m_spGrid->Row, lCol);

		// define "NorthWind Best Sellers Details" SQL query (SQL created with Access)
		_bstr_t strSQL = 
			"SELECT Customers.CompanyName As Customer, Orders.OrderDate As [Date], CCur([Order Details].[UnitPrice]*[Quantity]*(1-[Discount])/100)*100 AS [Value] "
			"FROM Customers INNER JOIN (Orders INNER JOIN [Order Details] ON Orders.OrderID = [Order Details].OrderID) ON Customers.CustomerID = Orders.CustomerID "
			"WHERE ((([Order Details].ProductID) = " + strProdID + ")) ORDER BY CCur([Order Details].[UnitPrice]*[Quantity]*(1-[Discount])/100)*100 DESC;";

		// create the detail recordset
		_RecordsetPtr rs;
		HRESULT hr = rs.CreateInstance(__uuidof(Recordset));
		if (FAILED(hr)) return 0;
		hr = rs->raw_Open(_variant_t(strSQL), _variant_t((IUnknown*)m_spConnMain),
                          adOpenUnspecified, adLockUnspecified, 0);
		if (FAILED(hr)) return 0;

		// bind grid to data source
		DataSourcePtr dataSource = rs->GetDataSource();
		if (dataSource == NULL) return 0;
		m_spGrid->DataMode = flexDMBound;
		m_spGrid->DataSource = dataSource;

		// done
		return 0;
	}


////////////////////////////////////////////////////////////////////////////////////////
//
// Event Handlers
//
////////////////////////////////////////////////////////////////////////////////////////
BEGIN_SINK_MAP(CDetailDlg)
	//Make sure the Event Handlers have __stdcall calling convention
	SINK_ENTRY(IDC_FGDETAIL, 0x6, OnBeforeRowColChange)
	SINK_ENTRY(IDC_FGDETAIL, 0x52, OnAfterDataRefresh)
	SINK_ENTRY(IDC_FGDETAIL, DISPID_KEYDOWN, OnKeyDown)
	SINK_ENTRY(IDC_FGDETAIL, 0x23, OnBeforeScrollTip)
END_SINK_MAP()

	VOID __stdcall OnBeforeRowColChange(LONG OldRow, LONG OldCol, LONG NewRow, LONG NewCol, VARIANT_BOOL * Cancel)
	{
		// update DB-like cursor glyph
		m_spGrid->PutTextMatrix(OldRow, 0, _bstr_t(""));
		m_spGrid->PutTextMatrix(NewRow, 0, _bstr_t("0"));
	}

	VOID __stdcall OnAfterDataRefresh()
	{
		// use imagelist to show db-like cursor glyph
		m_spGrid->PutColWidth(0, m_spGrid->GetRowHeight(0));
		m_spGrid->PutColImageList(0, (LONG)(HIMAGELIST)m_imgDBCursor);
		m_spGrid->PutColAlignment(0, flexAlignCenterCenter);
		m_spGrid->PutTextMatrix(1, 0, _bstr_t("0"));

		// freeze one column
		m_spGrid->FrozenCols = 1;
	}

	VOID __stdcall OnKeyDown(SHORT * KeyCode, SHORT Shift)
	{
		// hide when user pushes escape
        if (*KeyCode == VK_ESCAPE) {
			ShowWindow(SW_HIDE);
            m_spGrid->DataSource = NULL;
        }

	}

	VOID __stdcall OnBeforeScrollTip(LONG Row)
	{
		long lCol = m_spGrid->GetColIndex(_bstr_t("Customer"));
		_bstr_t strTip = " Customer: " + m_spGrid->GetTextMatrix(Row, lCol) + " ";
		m_spGrid->ScrollTipText = strTip;
	}


////////////////////////////////////////////////////////////////////////////////////////
//
// Member Variables
//
////////////////////////////////////////////////////////////////////////////////////////
	IVSFlexGridPtr	m_spGrid;
	CAxWindow		m_wndGrid;
	CImageList		m_imgDBCursor;

	// members initialized by caller
	IVSFlexGridPtr	m_spGridMain;
	_ConnectionPtr	m_spConnMain;
};

#endif //__DETAILDLG_H_
