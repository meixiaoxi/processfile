
// processfileDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "processfile.h"
#include "processfileDlg.h"
#include "afxdialogex.h"
#include<odbcinst.h> 
#include<afxdb.h>
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


// CprocessfileDlg 对话框



CprocessfileDlg::CprocessfileDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CprocessfileDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CprocessfileDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CprocessfileDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CprocessfileDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// CprocessfileDlg 消息处理程序

BOOL CprocessfileDlg::OnInitDialog()
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

	// TODO: 在此添加额外的初始化代码

HINSTANCE hInst=NULL;
hInst=AfxGetApp()->m_hInstance;
char path_buffer[_MAX_PATH];

	GetCurrentDirectory(_MAX_PATH,mDesFile);

	strcat_s(mDesFile,"\\desFile.csv");

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CprocessfileDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CprocessfileDlg::OnPaint()
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
HCURSOR CprocessfileDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

CString GetExcelDriver()
{
    char szBuf[2001];
    WORD cbBufMax = 2000;
    WORD cbBufOut;
    char *pszBuf = szBuf;
    CString sDriver;
    // 获取已安装驱动的名称(涵数在odbcinst.h里)
    if (!SQLGetInstalledDrivers((LPSTR)szBuf, cbBufMax, &cbBufOut))
    {
		sDriver = _T("");
		return _T(""); 
	}
    // 检索已安装的驱动是否有Excel...
    do
    {
        if (strstr(pszBuf, "Excel") != 0)
        {
            //发现 !
            sDriver = CString(pszBuf);
            break;
        }
        pszBuf = strchr(pszBuf, '\0') + 1;
    }
    while (pszBuf[1] != '\0');
    return sDriver;
}

#if 0
CString CprocessfileDlg::GetCellString(long irow, long icolumn)  
{  
     
    COleVariant vResult ;  
    CString str;  
    //字符串  
    if (already_preload_ == FALSE)  
    {  
        CRange range;  
       range.AttachDispatch(excel_current_range_.get_Item (COleVariant((long)irow),COleVariant((long)icolumn)).pdispVal, true);  
        vResult =range.get_Value2();  
        range.ReleaseDispatch();  
    }  
    //如果数据依据预先加载了  
    else  
    {  
        long read_address[2];  
        VARIANT val;  
        read_address[0] = irow;  
       read_address[1] = icolumn;  
        ole_safe_array_.GetElement(read_address, &val);  
        vResult = val;  
   }  
  
    if(vResult.vt == VT_BSTR)  
    {  
        str=vResult.bstrVal;  
    }  
    //整数  
    else if (vResult.vt==VT_INT)  
    {  
        str.Format("%d",vResult.pintVal);  
    }  
    //8字节的数字   
    else if (vResult.vt==VT_R8)       
    {  
        str.Format("%0.0f",vResult.dblVal);  
    }  
    //时间格式  
    else if(vResult.vt==VT_DATE)      
    {  
        SYSTEMTIME st;  
        VariantTimeToSystemTime(vResult.date, &st);  
       CTime tm(st);   
        str=tm.Format("%Y-%m-%d");  
  
    }  
    //单元格空的  
   else if(vResult.vt==VT_EMPTY)     
    {  
        str="";  
    }    
  
    return str;  
}  
#endif
char headFlag = 0;
BOOL CprocessfileDlg::mOperate()
{
//导入
    COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 
    if (!app.CreateDispatch(_T("Excel.Application")))
    {   
        this->MessageBox(_T("无法创建Excel应用！")); 
        return TRUE;  
    }
	if (!appMes.CreateDispatch(_T("Excel.Application")))
    {   
        this->MessageBox(_T("无法创建Excel应用！")); 
        return TRUE;  
    }
    books = app.get_Workbooks();
	booksMes= appMes.get_Workbooks();
    //打开Excel，其中pathname为Excel表的路径名  
    lpDispMes = booksMes.Open(_T("F:\\work\\processfile\\processfile\\Debug\\导出MES 原始档.xsl.xls"),covOptional ,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional);
    bookMes.AttachDispatch(lpDispMes); 

	lpDisp = books.Open(_T("F:\\work\\processfile\\processfile\\Debug\\all_keywords1207.csv"),covOptional ,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional);
    book.AttachDispatch(lpDisp);


    sheets = book.get_Worksheets(); 
    sheet = sheets.get_Item(COleVariant((short)1));

	sheetsMes = bookMes.get_Worksheets(); 
    sheetMes = sheetsMes.get_Item(COleVariant((short)1));

    //获得坐标为（A，1）的单元格 
    //range = sheet.get_Range(COleVariant(_T("A1")) ,COleVariant(_T("A1")));  

	long rows,rowsMes;
	long temp,tempMes,tempW;
	CString str,strMes,strWrite,strTemp;
     COleVariant vResult;
	 CStdioFile file;
	////获取sheet所使用的范围
	range = sheet.get_UsedRange();  
	tRange = range;
	range = range.get_Rows();
	rows = range.get_Count();

	rangeMes = sheetMes.get_UsedRange();  
	tRangeMes = rangeMes;
	rangeMes = rangeMes.get_Rows();
	rowsMes = rangeMes.get_Count();

	file.Open(mDesFile,CFile::modeCreate | CFile::modeNoTruncate | CFile::modeWrite);

	if(file == NULL)
		return false;

	file.SeekToEnd();
	if(headFlag == 0)
	{
		file.WriteString("箱唛,盒号,CID,SN\n");
		headFlag = 1;
	}
	for(tempMes = 2; tempMes <= rowsMes;tempMes++)
	{
		rangeMes.AttachDispatch(tRangeMes.get_Item (COleVariant((long)tempMes),COleVariant((long)3)).pdispVal, true);
		vResult =rangeMes.get_Value2();  
		strMes=vResult.bstrVal;  
		
		for(temp = 2; temp <= rows; temp++)
		{
			range.AttachDispatch(tRange.get_Item (COleVariant((long)temp),COleVariant((long)3)).pdispVal, true);
			vResult =range.get_Value2();  
			str=vResult.bstrVal;  
			if(str == strMes)
			{
				strWrite = _T("");
				for(tempW = 1;tempW < 3; tempW++)
				{
					if(tempW != 1)
						strWrite += _T(",");
					rangeMes.AttachDispatch(tRangeMes.get_Item (COleVariant((long)tempMes),COleVariant((long)tempW)).pdispVal, true);
					vResult =rangeMes.get_Value2();  
					strTemp=vResult.bstrVal;
					strWrite += strTemp; 
				}
				strWrite += _T(",");
				strWrite += str;
				strWrite += _T(",");
				range.AttachDispatch(tRange.get_Item (COleVariant((long)temp),COleVariant((long)4)).pdispVal, true);
				vResult =range.get_Value2();  
				str=vResult.bstrVal;  
				strWrite += str;

				strWrite += _T("\n");
				file.SeekToEnd();
		if(headFlag == 0)
		{
			file.WriteString("箱唛,盒号,CID,SN\n");
			headFlag = 1;
		}
				file.WriteString(strWrite);
				break;
			}
		}
	}

	file.Close();
	 
    app.Quit();
	appMes.Quit();
	return TRUE;
}
void CprocessfileDlg::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码
	  CDatabase dataMes,dataKeyWord;
	  CString sDriver;
	  CString sFileMes;
	  CString opFileMes;
	  CString sFileKeyWord;
	  CString opFileKeyWord;

	  GetDlgItem(IDC_BUTTON1)->EnableWindow(FALSE);

	  mOperate();
	  GetDlgItem(IDC_BUTTON1)->EnableWindow(TRUE);
#if 0
	  sFileMes = _T("导出MES 原始档.xsl.xls"); 
	  sFileKeyWord = _T("all_keywords1207.csv");
	  sDriver = GetExcelDriver();

		

     if (sDriver.IsEmpty())
	 {
        // 没有发现Excel驱动
        AfxMessageBox(_T("没有安装Excel驱动!"));
       return;
	 }

	  opFileMes.Format(_T("ODBC;DRIVER={%s};DSN='';DBQ=%s"), sDriver, sFileMes);
	  opFileKeyWord.Format(_T("ODBC;DRIVER={%s};DSN='';DBQ=%s"), sDriver, sFileKeyWord);

	  TRY
	  {
		dataMes.Open(NULL,false,false,opFileMes);
		dataMes.Open(NULL,false,false,opFileKeyWord);


		CRecordset recsetMes(&dataMes),recsetKeyWord(&dataKeyWord);
		CString sqlMes,sqlKeyWord;
		CString sSn;

		recsetMes.MoveNext();
		recsetKeyWord.MoveNext();

		 sqlMes = "SELECT SN" //设置索引顺序    
               "FROM [Sheet1$]" ;               
               "";

		 recsetMes.Open(CRecordset::forwardOnly,sqlMes,CRecordset::readOnly);
		 while(!recsetMes.IsEOF())
		 {
			 recsetMes.GetFieldValue(_T("SN"),sSn);
		 }
	  }
	  CATCH(CDBException, e)
     {
        // 数据库操作产生异常时...
        AfxMessageBox(_T("数据库错误: "));
     }
     END_CATCH;
 
#endif

}
