
// processfileDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "processfile.h"
#include "processfileDlg.h"
#include "afxdialogex.h"
#include<odbcinst.h> 
#include<afxdb.h>
#include "workthread.h"
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

	mFile[0] = '\0';
	mFileMes[0] = '\0';

	strcat_s(mFile,mDesFile);
	strcat_s(mFileMes,mDesFile);
	strcat_s(mDesFile,"\\怪兽充电出货SN&MAC对应表.csv");
	strcat_s(mFile,"\\all_keywords.csv");
	strcat_s(mFileMes,"\\导出.xsl.xls");
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
char headFlag = 0,index=0;
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
 //   lpDispMes = booksMes.Open(_T("F:\\work\\processfile\\processfile\\Debug\\导出MES 原始档.xsl.xls"),covOptional ,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional);
	lpDispMes = booksMes.Open(mFileMes,covOptional ,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional);
	bookMes.AttachDispatch(lpDispMes); 

	//lpDisp = books.Open(_T("F:\\work\\processfile\\processfile\\Debug\\all_keywords1207.csv"),covOptional ,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional);
    lpDisp = books.Open(mFile,covOptional ,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional);
	book.AttachDispatch(lpDisp);


    sheets = book.get_Worksheets(); 
    sheet = sheets.get_Item(COleVariant((short)1));
	rangeFind.AttachDispatch(sheet.get_Cells());
	rangeRow = sheet.get_Range(COleVariant(_T("C1")) ,COleVariant(_T("C1000")));

	sheetsMes = bookMes.get_Worksheets(); 
    sheetMes = sheetsMes.get_Item(COleVariant((short)1));

    //获得坐标为（A，1）的单元格 
    //range = sheet.get_Range(COleVariant(_T("A1")) ,COleVariant(_T("A1")));  

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


	rangeMes = sheetMes.get_UsedRange();  
	tRangeMes = rangeMes;
	rangeMes = rangeMes.get_Rows();
	rowsMes = rangeMes.get_Count();


	char tt[10],te[10];
	
	sprintf(tt,"C%d",rows);
	

	rangeRow = sheet.get_Range(COleVariant(_T("C1")) ,COleVariant(tt));

#if 0

	while(rowsMes > 5000)
	{
		memset(tt,10,'\0');
		memset(te,10,'\0');
		sprintf(tt,"A%d",tRowF);
		sprintf(te,"F%d",tRowE);
		m_info[index].RangeMes = sheetMes.get_Range(COleVariant(tt) ,COleVariant(te));
		m_info[index].RangeRow = rangeRow;
		m_info[index].RangeFind = rangeFind;
		m_info[index].tRange = tRange;
		m_info[index].cnt = 5000;
		tRowF +=5000;
		tRowE +=5000;
		rows-=5000;
		m_info[index].pThread = AfxBeginThread((AFX_THREADPROC)operateChild, &(m_info[index]),THREAD_PRIORITY_NORMAL,0,0,NULL);
		index++;
	}

#endif
	file.Open(mDesFile,CFile::modeCreate | CFile::modeNoTruncate | CFile::modeWrite);

	if(file == NULL)
		return false;

	file.SeekToEnd();
	if(headFlag == 0)
	{
		file.WriteString("箱唛,盒号,CID,扫描时间,SN\n");
		headFlag = 1;
	}
	strWrite = _T("");
	for(tempMes = 2; tempMes <= rowsMes;tempMes++)
	{		
		for(tempW = 1;tempW <=5; tempW++)
		{
			if(tempW == 4)
				continue;
			if(tempW != 1)
				strWrite += _T(",");
			rangeMes.AttachDispatch(tRangeMes.get_Item (COleVariant((long)tempMes),COleVariant((long)tempW)).pdispVal, true);
			vResult =rangeMes.get_Value2();  
			strTemp=vResult.bstrVal;
			strWrite += strTemp; 
			if(tempW ==3)
			{
				strMes = strTemp;	
			}
		}



		lpDispFind = rangeRow.Find(COleVariant(strMes), vOpt, xlValues, xlPart,xlByColumns, xlNext, xlIgnoreCase, vOpt,vOpt);
			if(lpDispFind)
			{
				CRange rTemp;
				rTemp = rangeFind;
				rTemp.AttachDispatch(lpDispFind);
				rTemp.Select();
				rTemp.Activate();
				mRow = rTemp.get_Row();
				mCol = rTemp.get_Column();
				strWrite += _T(",");
				range.AttachDispatch(tRange.get_Item (COleVariant((long)mRow),COleVariant((long)(mCol+1))).pdispVal, true);
				vResult =range.get_Value2();  
				str=vResult.bstrVal;  
				strWrite += str;
			}
			strWrite += _T("\n");

			if(tempFile++ > 50)
			{
				tempFile = 0;
				file.WriteString(strWrite);
				strWrite = _T("");
			}
	}
	if(tempFile > 0)
	{
		file.WriteString(strWrite);
	}
	AfxMessageBox("生成完毕");
	file.Close();
	 
	tRange.ReleaseDispatch();
	rangeMes.ReleaseDispatch();
	rangeFind.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheetMes.ReleaseDispatch();
	book.ReleaseDispatch();
	bookMes.ReleaseDispatch();

    app.Quit();
	appMes.Quit();
	app.ReleaseDispatch();
	appMes.ReleaseDispatch();
	return TRUE;
}




void CprocessfileDlg::mFindCidSN()
{
	char path_buffer[_MAX_PATH];
	CString filenames[1024],path;
	CFileFind finder;
	int count =0;
	BOOL working;



	//获取当前路径
	GetCurrentDirectory(_MAX_PATH,path_buffer);

	path = (CString)path_buffer;
	working = finder.FindFile(path + "\\*.txt");

	while (working)
	{
		working = finder.FindNextFile();
		if (finder.IsDots())
			continue;
		if (finder.IsDirectory())
		{
			//FindAllFile(finder.GetFilePath(), filenames, count);
		} 
		else 
		{
			CString filename = finder.GetFileName();
			filenames[count++] = filename;
		}
	}


	CStdioFile file,csvFile;
	CString readContnent;
	CString readBuffer[3];
	int posBuffer[3];
	const char *pos;
	char *fuck;
	char headFlag = 0;
	int totalNum = 0,tempNum =0,tempPos,statusFlag = 0xFF;
	// info file 
		/*		ch =  (LPCTSTR)mLogSavePath;
				len = mLogSavePath.GetLength();

				strncpy(mInfoPath,ch,len);
				*/
#if 0
	FILE *fp;

					  fopen_s(&fp,"result.txt","w");
				   if(fp ==NULL)
				   {
					   AfxMessageBox("save fail");
				   }
#endif
	csvFile.Open("result.csv",CFile::modeCreate | CFile::modeNoTruncate | CFile::modeWrite);
	while(count)
	{
		count--;
		pos = (LPCTSTR)filenames[count];
		if(strncmp(pos,"result.txt",10) == 0)
			continue;
		 if(file.Open(pos,CStdioFile::modeRead)==false)
       { 
#if 0
		   fclose(fp);
#endif
		   csvFile.Close();
            AfxMessageBox("打开文件失败");
       }
	
		// fprintf(fp,pos);
		 //fprintf(fp,"\r\n");
		 statusFlag = 0xFF;

		 while(file.ReadString(readContnent))//获取文件的长度，到文件末尾时返回false；
       {
		   tempPos = readContnent.Find("Cid = ");
		   if(tempPos >= 0)
		   {
			    statusFlag =0;
				readBuffer[statusFlag] = readContnent.Mid(tempPos+6,9);
		   }
		   tempPos = readContnent.Find("id=");
		   if(tempPos >= 0)
		   {
			   if(statusFlag != 0)
			   {
				   statusFlag = 0xFF;
			   }
			   else
			   {
				   statusFlag = 1;
				   readBuffer[statusFlag] = readContnent.Mid(tempPos+3,32);
			   }
		   }
		    tempPos = readContnent.Find("clear ok");
		   if(tempPos >= 0)
		   {
			   if(statusFlag != 1)
			   {
				   statusFlag = 0xFF;
			   }
			   else
			   {
				   statusFlag = 2;
				   statusFlag = 0xFF;
				   //save
					#if 0
				   pos = (LPCTSTR)readBuffer[0];
				   strncpy_s(path_buffer,pos,9);
				    pos = (LPCTSTR)readBuffer[1];
					strncpy(path_buffer+10,pos,32);

				   path_buffer[9] = ' ';
				   path_buffer[42] = '\0';

				   fprintf(fp,path_buffer);
				   fprintf(fp,"\n");
				#endif
				   totalNum++;
				   if(totalNum >=540)
					   totalNum++;
				   if(headFlag == 0)
				   {
					   csvFile.WriteString("CID,SN\n");
					   headFlag = 1;
				   }
				   csvFile.WriteString(readBuffer[0]+_T(",")+readBuffer[1]+_T("\n"));
			   }
		   }
       }

		 file.Close();
	}

	//AfxMessageBox("生成完毕");
	MessageBox("生成完毕","",MB_OK);
#if 0
	fclose(fp);
#endif
	csvFile.Close();
}

void CprocessfileDlg::OnCancel()
{
		 PostMessage(WM_QUIT,0,0);//最常用
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
		
	  //mFindCidSN();
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
