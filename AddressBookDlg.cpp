// AddressBookDlg.cpp : implementation file
//

#include "stdafx.h"
#include "AddressBookDlg.h"
#include "InternalAddressBookSearchWnd.h"


CString CAddressBookDlg::m_sInternal;
CString CAddressBookDlg::m_sExternal;
CString CAddressBookDlg::m_sOk;
CString CAddressBookDlg::m_sCancel;

// CMCAddressBookDlg dialog


IMPLEMENT_DYNAMIC(CAddressBookDlg, CDialog)

CAddressBookDlg::CAddressBookDlg(CWnd* pParent /*=NULL*/, ADDRBOOKMODE::AddressBookMode nMode, BOOL bLoadDLFavorites, BOOL bLoadPool, BOOL bShowExternalTab, BOOL bDefaultMethodOnly, RECIPIENT_DESTINATION eDestinationAddr)
	: CDialogEx(CAddressBookDlg::IDD, pParent),
	m_marginPix(6),
	m_withinPix(5),
	m_bShowExternalTab(bShowExternalTab),
	m_bShowInternalTab(TRUE),
	m_bBypassInternalEmail(FALSE),
	m_wndInternalSearch(nMode, bShowExternalTab, bLoadDLFavorites, bLoadPool, eDestinationAddr)
{
	CCrmVerifierPtr pCrmVerifier = CApplication::GetCrmVerifier();
	if (pCrmVerifier == NULL)
		CPS_ERROR_RETURN2("CAddressBookDlg::CAddressBookDlg - pCrmVerifier returned from CApplication::GetCrmVerifier() is NULL", "pvInbox");

	if (pCrmVerifier->IsExternalSearchMpageAvailable() == TRUE)
	{
		m_externalSearch.reset(new CExternalAddressBookSearchMpage(NULL, eDestinationAddr, bDefaultMethodOnly));
	}
	else
	{
		m_externalSearch.reset(new CExternalAddressBookSearchWnd(NULL, eDestinationAddr, bDefaultMethodOnly));
	}
	m_wndExternalSearch = std::dynamic_pointer_cast<CWnd>(m_externalSearch);

	m_font.DeleteObject();

	LOGFONT lfIcon;
	::ZeroMemory(&lfIcon, sizeof(LOGFONT));

	VERIFY(::SystemParametersInfo(SPI_GETICONTITLELOGFONT,sizeof(LOGFONT), &lfIcon, 0));

	lfIcon.lfWeight = FW_NORMAL;
	m_font.CreateFontIndirect(&lfIcon);

	m_hBrush = CreateSolidBrush(GetSchemeColor(gDefaultBackgroundColor));
}

CAddressBookDlg::~CAddressBookDlg()
{
	m_font.DeleteObject();
	DeleteObject(m_hBrush);
}


BEGIN_MESSAGE_MAP(CAddressBookDlg, CDialogEx)

	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_WM_PAINT()
	ON_WM_CTLCOLOR()
	ON_MESSAGE(WM_RECIPIENT_SELECTED, OnMsgToEnableOKBtn)
	ON_MESSAGE(WM_CNG_FAV_INT, OnMsgFavChngInt)
	ON_MESSAGE(WM_CNG_FAV_EXT, OnMsgFavChngExt)
	ON_MESSAGE(WM_SET_FOCUS_INTERNAL_TAB, OnSetFocusInternalTab)
	ON_BN_CLICKED(IDC_ADDRESSS_BOOK_OK, OnOK)

END_MESSAGE_MAP()

// CMCAddressBookDlg message handlers
int CAddressBookDlg::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CDialogEx::OnCreate(lpCreateStruct) == -1)
		return -1;

	if(m_sExternal.IsEmpty() == TRUE)
		m_sExternal.LoadString(IDS_STATUS_EXTERNAL);

	if(m_sInternal.IsEmpty() == TRUE)
		m_sInternal.LoadStringA(IDS_STATUS_INTERNAL);

	if(m_sOk.IsEmpty() == TRUE)
		m_sOk.LoadString(IDS_OK);

	if(m_sCancel.IsEmpty() == TRUE)
		m_sCancel.LoadStringA(IDS_CANCEL);

	CRect rect(0,0,0,0);
	SetupTabControl();

	if (m_wndInternalSearch.Create(NULL, "Internal Address", WS_CHILD | WS_VISIBLE, rect, &m_tabControl, IDC_INTERNAL_ADDRESS_BOOK) == FALSE)
		CPS_ERROR_RETURN("CAddressBookDlg::OnCreate - Failed to Create Internal Search window", "PVInbox", -1);
	m_wndInternalSearch.ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT);

	m_wndInternalSearch.SetHandle(this->GetSafeHwnd());
	m_tabControl.InsertItem(0, m_sInternal, m_wndInternalSearch.GetSafeHwnd());
	if (m_bShowInternalTab == FALSE)
	{
		m_tabControl.GetItem(0)->SetVisible(FALSE);
	}

	if(m_bShowExternalTab == TRUE)
	{
		CUserPtr pUserPtr = CApplication::GetUser();
		if (pUserPtr == NULL)
			CPS_ERROR_RETURN("CAddressBookDlg::OnCreate - GetUser failed", "PVInbox",-1);

		CCrmVerifierPtr pCrmVerifier = CApplication::GetCrmVerifier();
		if (pCrmVerifier == NULL)
			CPS_ERROR_RETURN("CAddressBookDlg::OnCreate - pCrmVerifier returned from CApplication::GetCrmVerifier() is NULL", "pvInbox", FALSE);

		if ((m_bBypassInternalEmail == TRUE || pUserPtr->CanUserSecureMessage()))
		{
			if (m_wndExternalSearch->Create(NULL, "External Address", WS_CHILD | WS_VISIBLE, rect, &m_tabControl, IDC_EXTERNAL_ADDRESS_BOOK_MPAGE) == FALSE)
				CPS_ERROR_RETURN("CAddressBookDlg::OnCreate - Failed to Create External Search window", "PVInbox", -1);
			m_wndExternalSearch->ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT);

			m_externalSearch->SetExternalHandle(this->GetSafeHwnd());
			m_tabControl.InsertItem(1, m_sExternal, m_wndExternalSearch->GetSafeHwnd());
			if (m_bShowInternalTab == FALSE)
				m_tabControl.GetItem(1)->Select();
		}
		else if (m_bShowInternalTab == FALSE)
		{
			CPS_ERROR_RETURN("CAddressBookDlg::OnCreate - Failed to show External Search window and Internal Search window", "PVInbox", -1);
		}
	}

	if (m_btnOK.Create(m_sOk, WS_CHILD | WS_VISIBLE | WS_TABSTOP, rect, this, IDC_ADDRESSS_BOOK_OK) == FALSE)
		CPS_ERROR_RETURN("CAddressBookDlg::OnCreate - Failed to Create Ok button", "PVInbox", -1);

	if (m_btnCancel.Create(m_sCancel, WS_CHILD | WS_VISIBLE | WS_TABSTOP, rect, this, IDCANCEL) == FALSE)
		CPS_ERROR_RETURN("CAddressBookDlg::OnCreate - Failed to Create Cancel button", "PVInbox", -1);

	m_tabControl.Reposition();
			
	if(&m_font)
	{
		m_btnOK.SetFont(&m_font);
		m_btnCancel.SetFont(&m_font);	
	}

	m_sizeUtility.SetParent(GetSafeHwnd());
	
	return 0;
}


void CAddressBookDlg::OnSize(UINT nType, int cx, int cy)
{
	CDialogEx::OnSize(nType, cx, cy);

	CRect rect(0,0,0,0);
	int iRowHeight = m_sizeUtility.StaticHeight() + m_sizeUtility.ConvertY(1) + m_sizeUtility.StaticOffset();
	
	m_tabControl.MoveWindow(m_marginPix * 2, m_marginPix * 2, cx - m_marginPix  * 4, cy - iRowHeight - m_marginPix * 4 - m_withinPix * 2);

	//if block added as part of CR 1-11071095661
	//When resizing external MPage tab MoveWindow will call discernoutputviewer's onsize function which resizes external tab asynchronously.
	//Updating the window is required after external tab is resized by DOV to paint external tab properly 
	if (m_bShowExternalTab == TRUE)
	{
		CCrmVerifierPtr pCrmVerifier = CApplication::GetCrmVerifier();
		if(pCrmVerifier == NULL)
			CPS_ERROR_RETURN2("CAddressBookDlg::OnSize returned from CApplication::GetCrmVerifier() is NULL", "pvInbox");
		if ( pCrmVerifier->IsExternalSearchMpageAvailable())
		{
			::PostMessage(GetSafeHwnd(), WM_UPDATE_WND, NULL, NULL);
		}
	}

	m_tabControl.GetWindowRect(&rect);
	ScreenToClient(&rect);

	m_btnCancel.MoveWindow(cx - 100 - m_marginPix * 2, cy - iRowHeight - m_marginPix * 2, 100, iRowHeight);
	m_btnCancel.GetWindowRect(&rect);
	ScreenToClient(&rect);

	m_btnOK.MoveWindow(rect.left - m_withinPix * 2 - 100, rect.top, 100, iRowHeight);
	m_btnOK.GetWindowRect(&rect);
	ScreenToClient(&rect);

	Invalidate(TRUE);
}


void CAddressBookDlg::SetupTabControl()
{
	//Create TabControl
	if(m_tabControl.Create(WS_CHILD|WS_VISIBLE, CRect(0, 0, 0, 0), this, IDC_ADDRSEARCHTABCTRL) == FALSE)
		CPS_ERROR_RETURN2("CAddressBookDlg::SetupTabControl - Failed to Create External Search Tab", "PVInbox");


	m_tabControl.ModifyStyleEx(0, WS_EX_CONTROLPARENT);
	m_tabControl.GetPaintManager()->m_clientFrame = xtpTabFrameBorder;

	m_tabControl.GetPaintManager()->SetAppearance(xtpTabAppearancePropertyPage);
	m_tabControl.GetPaintManager()->SetColor(xtpTabColorWinNative);
	m_tabControl.GetPaintManager()->m_rcClientMargin.SetRect(0,1,0,0);
	m_tabControl.GetPaintManager()->m_rcButtonMargin.SetRect(0,1,0,1);

	m_tabControl.GetPaintManager()->m_bShowIcons = FALSE;
	m_tabControl.GetPaintManager()->DisableLunaColors(TRUE);
	m_tabControl.GetPaintManager()->m_bBoldSelected = TRUE;
	m_tabControl.SetLayoutStyle(xtpTabLayoutAutoSize);
}

int CAddressBookDlg::GetTextWidth(CWnd *wnd, int nMinWidth)
{
	if(!(wnd->GetStyle() & WS_VISIBLE))
		return nMinWidth;

	CString sControlText = "";
	wnd->GetWindowText(sControlText);
	if(sControlText.IsEmpty())
		return nMinWidth;
	
	CDC* pDC = wnd->GetDC();
	if (pDC == NULL)
		return nMinWidth;

	CFont* pOldFont = pDC->SelectObject(wnd->GetFont());
	
	int nLineWidth(0);
	char* pContext = NULL;
	CString lineText;

	//look at the width of each line of text in the control (example: Comments:\nLimit 255)
	lineText =  strtok_s(sControlText.GetBuffer(), "\n", &pContext );
	do
	{
		nLineWidth = pDC->GetTextExtent(lineText).cx;
		if(nLineWidth > nMinWidth)
			nMinWidth = nLineWidth;

		lineText =  strtok_s(NULL, "\n", &pContext );
	} while(lineText.IsEmpty() == FALSE);

	pDC->SelectObject(pOldFont);

	return nMinWidth;
}

void CAddressBookDlg::OnPaint()
{
	CPaintDC dc(this);

	// paint the back of the window white and keep the rest of the window gray
	//  we're going to grab the rectangles of the window and the Filter groupbox
	//  and then XOR the two rectanges to create a region that we'll paint white
	CBrush WindowBrush(::GetSchemeColor(gDefaultBackgroundColor));
	
	CRect WindowRect;
	GetClientRect(&WindowRect);	

	CBrush bgBrush(GetSchemeColor(gDefaultBackgroundColor));
	CBrush* pBrush = dc.SelectObject(&bgBrush);
	dc.FillRect(&WindowRect, &bgBrush);

	dc.SelectObject(pBrush);
}

HBRUSH CAddressBookDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);

	pDC->SetBkColor(GetSchemeColor(gDefaultBackgroundColor));
	hbr =  ::GetSchemeBrush(gBodyBackgroundBrush);
	
	return hbr;
}


BOOL CAddressBookDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	
	BOOL bFocusExternalTab(FALSE);

	
	CUserPtr pUserPtr = CApplication::GetUser();
	if(pUserPtr == NULL)
		CPS_ERROR_RETURN("CExternalAddrController::SearchForRecipients - GetUser failed", "PVInbox",FALSE);
	
	if(m_sPreSearchName.GetAt(0) == '/'||m_sPreSearchName.GetAt(0) == '.')
	{
		if(m_sPreSearchName.GetAt(1) == 'e' || m_sPreSearchName.GetAt(1) == 'E')
		{
			m_sPreSearchName = m_sPreSearchName.Mid(2);
			if (m_bShowExternalTab == TRUE)
				bFocusExternalTab = TRUE;
		}
	}
	
	if(bFocusExternalTab == TRUE || m_bShowInternalTab == FALSE) 
	{
		m_tabControl.SetCurSel(1);
		m_externalSearch->SetSearchName(m_sPreSearchName);
		m_tabControl.Reposition();
	}
	else
	{
		m_wndInternalSearch.SetSearchName(m_sPreSearchName);
	}
	

	m_wndInternalSearch.InitializeWnd();
	m_wndInternalSearch.GetRecipients(m_vecRecipient, m_vecDLRecipient);

	//When the search is not supported (searching for pools or Dls for prov letter workflow) 
	//or when we search for a recipient with full name in the TO Field , if the name is there in Address Book, the following condition will get succeed.
	//When the recipient can be resolved to one, then return from here.
	//This check has to be done before we check for the external tab condition. 
	if (m_wndInternalSearch.IsSearchNotSupported() || m_vecRecipient.size() == 1 || m_vecDLRecipient.size() == 1)
	{
		CDialogEx::OnOK();
		return TRUE;
	}
	if (m_bShowExternalTab == TRUE || m_wndExternalSearch->GetSafeHwnd() != NULL)
	{
		m_externalSearch->InitializeWnd();
		m_externalSearch->GetRecipients(m_vecRecipient);
		//To handle the automatic resolving to one external user.
		if (m_vecRecipient.size() == 1)
			CDialogEx::OnOK();
	}

	m_btnOK.EnableWindow(FALSE);
	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

LRESULT CAddressBookDlg::OnMsgFavChngExt(WPARAM wParam, LPARAM lParam)
{
	m_wndInternalSearch.FavoritesUpdatedFromExternal();
	return TRUE;
}

LRESULT CAddressBookDlg::OnMsgFavChngInt(WPARAM wParam, LPARAM lParam)
{
	m_externalSearch->FavoritesChangedFromInternal();
	return TRUE;
}

//Handler for WM_SET_FOCUS_INTERNAL_TAB message
//purpose: To set the focus to internal tab search box once external MPage tab finishes loading
LRESULT CAddressBookDlg::OnSetFocusInternalTab(WPARAM wParam, LPARAM lParam)
{
	if (m_tabControl.GetItem(0)->IsSelected()==TRUE)
	{
		m_wndInternalSearch.OnSetfocusEditName();	
	}	
	return TRUE;
}

LRESULT CAddressBookDlg::OnMsgToEnableOKBtn(WPARAM wParam, LPARAM lParam)
{
	if(wParam== TRUE)
	{
		m_btnOK.EnableWindow(TRUE);
	}
	else
	{
		std::vector<CAddrPerson*>			vecRecipient;
		std::vector<CDLRecipient*>			vecDLRecipient;

		// If the call came from the internal tab to disable the OK button then don't call GetRecipients on the internal tab
		if (lParam != (LPARAM)eInternalTab)
			m_wndInternalSearch.GetRecipients(vecRecipient, vecDLRecipient);
		
		// If the call came from the external tab to disable the OK button then don't call GetRecipients on the external tab
		if (lParam != (LPARAM)eExternalTab)
			m_externalSearch->GetRecipients(vecRecipient);

		if(vecRecipient.size()==0 && m_vecDLRecipient.size() == 0)
			m_btnOK.EnableWindow(FALSE);
	}
	
	return TRUE;
}

void CAddressBookDlg::OnOK()
{
	m_wndInternalSearch.GetRecipients(m_vecRecipient, m_vecDLRecipient);
	m_externalSearch->GetRecipients(m_vecRecipient);

	CDialogEx::OnOK();
}

void CAddressBookDlg::OnCancel()
{
	CTimeKeeper::ResumeTimers();
	PostNcDestroy();
	CDialogEx::OnCancel();
}

BOOL CAddressBookDlg::HasPrsnlAddrBookChanged()
{
	return m_wndInternalSearch.HasPrsnlAddrBookChanged() || m_externalSearch->HasPrsnlAddrBookChanged();
}

void CAddressBookDlg::GetRecipients(std::vector<CAddrPerson*>& vecRecipient, std::vector<CDLRecipient*>& vecDLRecipient)
{
	std::vector<CAddrPerson*>::iterator itRecipient = m_vecRecipient.begin();
	for(; itRecipient != m_vecRecipient.end(); itRecipient++)
	{
		vecRecipient.push_back(*itRecipient);
	}

	std::vector<CDLRecipient*>::iterator itDLRecipient = m_vecDLRecipient.begin();
	for(; itDLRecipient != m_vecDLRecipient.end(); itDLRecipient++)
	{
		vecDLRecipient.push_back(*itDLRecipient);
	}
}

void CAddressBookDlg::GetRecipients(std::vector<CAddrPerson*>& vecRecipient)
{
	std::vector<CAddrPerson*>::iterator itRecipient = m_vecRecipient.begin();
	for (; itRecipient != m_vecRecipient.end(); itRecipient++)
	{
		vecRecipient.push_back(*itRecipient);
	}
}

BOOL CAddressBookDlg::PreTranslateMessage(MSG* pMsg)
{
	if (pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_RETURN)
	{
		CWnd *pFocusWnd = GetFocus();
		if (pFocusWnd != &m_btnOK && pFocusWnd != &m_btnCancel)
			return TRUE;
	}
	return CDialogEx::PreTranslateMessage(pMsg);
}
