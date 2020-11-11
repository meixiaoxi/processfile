
// findsnDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "findsn.h"
#include "findsnDlg.h"
#include "afxdialogex.h"
#include "mymsg.h"
#include "util.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CfindsnDlg 对话框



CfindsnDlg::CfindsnDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CfindsnDlg::IDD, pParent)
{
	m_strEdit1 = _T("");
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CfindsnDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);

	DDX_Control(pDX, IDC_RICHEDIT22, m_ctrlRedit);

	DDX_Control(pDX, IDC_EDIT1, m_ctlEdit1);
	DDX_Control(pDX, IDC_EDIT2, m_ctlEdit2);

	DDX_Text(pDX, IDC_EDIT1, m_strEdit1);
	DDX_Text(pDX, IDC_EDIT2, m_strEdit2);
	DDX_Text(pDX,IDC_RICHEDIT22, m_strRedit);
}

BEGIN_MESSAGE_MAP(CfindsnDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_MESSAGE(WM_USER_NOTIFY,OnUserMsg)
	ON_BN_CLICKED(IDOK, &CfindsnDlg::OnBnClickedOk)
	//ON_WM_SIZE()
END_MESSAGE_MAP()


// CfindsnDlg 消息处理程序

BOOL CfindsnDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。
	
	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	::EnableMenuItem(::GetSystemMenu(this->m_hWnd, false), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);//forbid close

	// TODO: 在此添加额外的初始化代码





	m_info.hWnd = m_hWnd;
	m_info.pThread = AfxBeginThread(ThreadFunc, &m_info);
	m_info.isExcelAppCreate = 0;
	m_info.isSaveFileCreate = 0;

	GetCurrentDirectory(_MAX_PATH,workPath);
	snSrcFile[0]='\0';

	strcat_s(snSrcFile,workPath);
//	strcat_s(mFileMes,workPath);
	strcat_s(snSrcFile,"\\source.xlsx");
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CfindsnDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CfindsnDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CfindsnDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

BOOL CfindsnDlg::PreTranslateMessage(MSG* pMsg)
{
	char strTemp[10];
	if (pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_RETURN)
	{
			UpdateData();
			if (m_strEdit1.GetLength() == 9)
			{
				strcpy(strTemp, (LPCSTR)(m_strEdit1));
				if(strTemp[0] < 'A' || strTemp[0] > 'Z')
				{
					MessageBox("序列号首字母异常");
					m_strEdit1 = "";
					UpdateData(FALSE);
				}
				else
				{
					strcpy(m_info.num, (LPCSTR)(m_strEdit1));
					mOperate();
					m_strEdit1 = "";
					UpdateData(FALSE);
				}
			}
			else
			{
				#ifdef MANUAL_UNBIND_TEST
				MessageBox("证书长度为32位");
				#else
				MessageBox("Cid长度为9位");
				#endif
				m_strEdit1 = "";
				UpdateData(FALSE);
			}
			return TRUE;
	}
	
		// keep return/esc key
	if (pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
		return TRUE;
	
	return CDialog::PreTranslateMessage(pMsg);
}


#define xlComments COleVariant( -4144L )
   #define xlFormulas COleVariant( -4123L ) // will find value in any cell
   #define xlValues COleVariant( -4163L ) // ignores hidden cells

   //LookAt
   #define xlWhole COleVariant( 1L ) // whole word search
   #define xlPart COleVariant( 2L ) // partial word search

   //SearchOrder (vOpt works here)
   #define xlByRows COleVariant( 1L )
   #define xlByColumns COleVariant( 2L )

   //SearchDirection (required but usually has no effect)
   #define xlNext 1L
   #define xlPrev 2L

   // MatchCase
   #define xlMatchCase COleVariant( 1L )
   #define xlIgnoreCase COleVariant( 0L )

   // MatchByte
   // ignored, use vOpt

   _variant_t vOpt(DISP_E_PARAMNOTFOUND, VT_ERROR);

BOOL CfindsnDlg::mOperate()
{
  COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 
   char buf[500];
	if(m_info.isExcelAppCreate == 0)
   {
	   SendMessage(WM_USER_NOTIFY,WP_START_EXCEL_APP,0);
		 if (!app.CreateDispatch(_T("Excel.Application")))
    {   
        this->MessageBox(_T("无法创建Excel应用！")); 
        return TRUE;  
    }
	   books = app.get_Workbooks();

	try{
	lpDisp = books.Open(snSrcFile,covOptional ,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional);
	ASSERT(lpDisp);
	book.AttachDispatch(lpDisp); 
	}
	catch(CException *e)
	{
				book.ReleaseDispatch();


    app.Quit();
	
	app.ReleaseDispatch();

	SendMessage(WM_USER_NOTIFY,WP_FAIL_OPEN_SN_SOURCE_FILE,0);
		//MessageBox(0,TEXT("Could not open workbork."), MB_OK | MB_ICONERROR);
		return 0;
	}

	sheets = book.get_Worksheets(); 
    sheet = sheets.get_Item(COleVariant((short)1));
		long rows,rowsMes,tRowF=1,tRowE=5000;
	long temp,tempMes,tempW,tempFile = 0;
	long mRow,mCol;
	CString str,strMes,strWrite,strTemp;
     COleVariant vResult;
	 CStdioFile file;
	////获取sheet所使用的范围
	range = sheet.get_UsedRange();  
	tRange = range;
	range = range.get_Rows();
	rows = range.get_Count();


	range = sheet.get_UsedRange();  
	tRange = range;
	range = range.get_Rows();
	rowsMes = range.get_Count();


	char tt[10],te[10];
	
	sprintf(tt,"A%d",rows);
	

	rangeRow = sheet.get_Range(COleVariant(_T("A1")) ,COleVariant(tt));

	}

	m_info.isExcelAppCreate = 1;

	if(m_info.isSaveFileCreate == 0)
	{
		CTime tm;
		CString tfile_pass,tfile_fail,tfile_log;
		int len;
		tm=CTime::GetCurrentTime();

		tfile_log = tm.Format(_T("%Y-%m-%d-%H;%M;%S"))+_T(".txt");
		tfile_pass = tm.Format(_T("%Y-%m-%d-%H;%M;%S"))+_T("_PASS")+_T(".csv");
		tfile_fail = tm.Format(_T("%Y-%m-%d-%H;%M;%S"))+_T("_FAIL")+_T(".csv");
		const char *ch;
		// info file 
		ch =  (LPCTSTR)tfile_pass;
		len = tfile_pass.GetLength();
		strncpy(mSaveFilePassName,ch,len);
		mSaveFilePassName[len] = '\0';

		ch =  (LPCTSTR)tfile_fail;
		len = tfile_fail.GetLength();
		strncpy(mSaveFileFailName,ch,len);
		mSaveFileFailName[len] = '\0';

		ch =  (LPCTSTR)tfile_log;
		len = tfile_log.GetLength();
		strncpy(mLogFileName,ch,len);
		mLogFileName[len] = '\0';
	}

	m_info.isSaveFileCreate = 1;

	SendMessage(WM_USER_NOTIFY,WP_START_TEST,0);

	time_t ltime;
	CTime tm;
	time(&ltime);
	sprintf(buf, "start test\r\n%s\r\n扫入条码：%s\r\n", ctime(&ltime),m_info.num);

	SendMessage(WM_USER_NOTIFY,WP_PRINT_LOG_STR,(LPARAM)buf);
	

	lpDispFind = rangeRow.Find(COleVariant(m_info.num), vOpt, xlValues, xlPart,xlByColumns, xlNext, xlIgnoreCase, vOpt,vOpt);

	if(lpDispFind)
	{
		SendMessage(WM_USER_NOTIFY,WP_PRODUCT_IN_WARRANTY,0);
		SendMessage(WM_USER_NOTIFY,WP_SAVE_INFO_FILE_PASS,(LPARAM)m_info.num);
		sprintf(buf, "\r\n产品在保修期\r\n\r\n");

		SendMessage(WM_USER_NOTIFY,WP_PRINT_LOG_STR,(LPARAM)buf);
	}
	else
	{
		SendMessage(WM_USER_NOTIFY,WP_PRODUCT_EXPIRED_WARRANTY,0);
		SendMessage(WM_USER_NOTIFY,WP_SAVE_INFO_FILE_FAIL,(LPARAM)(m_info.num));
		sprintf(buf, "\r\n产品过保\r\n\r\n");

		SendMessage(WM_USER_NOTIFY,WP_PRINT_LOG_STR,(LPARAM)buf);
		MessageBox(_T("产品过保修期"),NULL, MB_ICONERROR);
	}


	
}

LRESULT CfindsnDlg::OnUserMsg(WPARAM wParam, LPARAM lParam) 
{
	CString strText;
	char tmpBuf[512];

	switch (wParam)
	{
		case WP_START_TEST:
			m_strRedit = "开始测试";
			m_ctrlRedit.SetBackgroundColor(FALSE, COLOR_YELLOW);
			UpdateData(FALSE);
			break;
		case WP_PRODUCT_EXPIRED_WARRANTY:
			m_strRedit = "过保修期";
			m_ctrlRedit.SetBackgroundColor(FALSE, COLOR_RED);
			UpdateData(FALSE);
			break;
		case WP_PRODUCT_IN_WARRANTY:
			m_strRedit = "可维修";
			m_ctrlRedit.SetBackgroundColor(FALSE, COLOR_GREEN);
			UpdateData(FALSE);
			break;
		case WP_FAIL_OPEN_SN_SOURCE_FILE:
			m_strRedit = "打开数据源文件失败";
			m_ctrlRedit.SetBackgroundColor(FALSE, COLOR_RED);
			UpdateData(FALSE);
			break;
		case WP_SAVE_INFO_FILE_PASS:
			CUtil::SaveInfoFile((char*)lParam,mSaveFilePassName);
			break;
		case WP_SAVE_INFO_FILE_FAIL:
			CUtil::SaveInfoFile((char*)lParam,mSaveFileFailName);
			break;
		case WP_PRINT_LOG_STR:
			// log 超过10k就取后5k
			//if (m_strEdit2.GetLength() > 10240)
			//	m_strEdit2 = m_strEdit2.Right(5000);

			m_strEdit2 += (char*)lParam;
			UpdateData(FALSE);

			{// auto scroll
				int count = m_ctlEdit2.GetLineCount();
				int first = m_ctlEdit2.GetFirstVisibleLine();
				if (count - first > 7)
				m_ctlEdit2.LineScroll(count - first - 7, 0);
			}
	#ifndef MANUAL_UNBIND_TEST
			CUtil::lymSaveLog((char*)lParam,mLogFileName);
	#endif
		break;
		case WP_START_EXCEL_APP:
			m_strRedit = "正在启动Excel APP...";
			m_ctrlRedit.SetBackgroundColor(FALSE, COLOR_YELLOW);
			UpdateData(FALSE);
			break;
	}

	return 0;
}


void CfindsnDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	CDialogEx::OnOK();

	if(m_info.isExcelAppCreate)
	{
		tRange.ReleaseDispatch();
	
	rangeFind.ReleaseDispatch();
	sheet.ReleaseDispatch();
	
	book.ReleaseDispatch();


    app.Quit();
	
	app.ReleaseDispatch();
	}
}


void CfindsnDlg::OnSize(UINT nType, int cx, int cy)
{
	CDialogEx::OnSize(nType, cx, cy);

	// TODO: 在此处添加消息处理程序代码
	 if(nType==SIZE_RESTORED||nType==SIZE_MAXIMIZED)
	{
		resize();
	}
}
void CfindsnDlg::resize()
{
float fsp[2];
POINT Newp; //获取现在对话框的大小
CRect recta;    
GetClientRect(&recta);     //取客户区大小  
Newp.x=recta.right-recta.left;
Newp.y=recta.bottom-recta.top;
fsp[0]=(float)Newp.x/Old.x;
fsp[1]=(float)Newp.y/Old.y;
CRect Rect;
int woc;
CPoint OldTLPoint,TLPoint; //左上角
CPoint OldBRPoint,BRPoint; //右下角
HWND  hwndChild=::GetWindow(m_hWnd,GW_CHILD);  //列出所有控件  
while(hwndChild)    
{    
woc=::GetDlgCtrlID(hwndChild);//取得ID
GetDlgItem(woc)->GetWindowRect(Rect);  
ScreenToClient(Rect);  
OldTLPoint = Rect.TopLeft();  
TLPoint.x = long(OldTLPoint.x*fsp[0]);  
TLPoint.y = long(OldTLPoint.y*fsp[1]);  
OldBRPoint = Rect.BottomRight();  
BRPoint.x = long(OldBRPoint.x *fsp[0]);  
BRPoint.y = long(OldBRPoint.y *fsp[1]);  
Rect.SetRect(TLPoint,BRPoint);  
GetDlgItem(woc)->MoveWindow(Rect,TRUE);
hwndChild=::GetWindow(hwndChild, GW_HWNDNEXT);    
}
Old=Newp;

}