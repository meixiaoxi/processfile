
// processfileDlg.cpp : ʵ���ļ�
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


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CprocessfileDlg �Ի���



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


// CprocessfileDlg ��Ϣ�������

BOOL CprocessfileDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������

HINSTANCE hInst=NULL;
hInst=AfxGetApp()->m_hInstance;
char path_buffer[_MAX_PATH];

	GetCurrentDirectory(_MAX_PATH,mDesFile);

	strcat_s(mDesFile,"\\desFile.csv");

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CprocessfileDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
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
    // ��ȡ�Ѱ�װ����������(������odbcinst.h��)
    if (!SQLGetInstalledDrivers((LPSTR)szBuf, cbBufMax, &cbBufOut))
    {
		sDriver = _T("");
		return _T(""); 
	}
    // �����Ѱ�װ�������Ƿ���Excel...
    do
    {
        if (strstr(pszBuf, "Excel") != 0)
        {
            //���� !
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
    //�ַ���  
    if (already_preload_ == FALSE)  
    {  
        CRange range;  
       range.AttachDispatch(excel_current_range_.get_Item (COleVariant((long)irow),COleVariant((long)icolumn)).pdispVal, true);  
        vResult =range.get_Value2();  
        range.ReleaseDispatch();  
    }  
    //�����������Ԥ�ȼ�����  
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
    //����  
    else if (vResult.vt==VT_INT)  
    {  
        str.Format("%d",vResult.pintVal);  
    }  
    //8�ֽڵ�����   
    else if (vResult.vt==VT_R8)       
    {  
        str.Format("%0.0f",vResult.dblVal);  
    }  
    //ʱ���ʽ  
    else if(vResult.vt==VT_DATE)      
    {  
        SYSTEMTIME st;  
        VariantTimeToSystemTime(vResult.date, &st);  
       CTime tm(st);   
        str=tm.Format("%Y-%m-%d");  
  
    }  
    //��Ԫ��յ�  
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
//����
    COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 
    if (!app.CreateDispatch(_T("Excel.Application")))
    {   
        this->MessageBox(_T("�޷�����ExcelӦ�ã�")); 
        return TRUE;  
    }
	if (!appMes.CreateDispatch(_T("Excel.Application")))
    {   
        this->MessageBox(_T("�޷�����ExcelӦ�ã�")); 
        return TRUE;  
    }
    books = app.get_Workbooks();
	booksMes= appMes.get_Workbooks();
    //��Excel������pathnameΪExcel���·����  
    lpDispMes = booksMes.Open(_T("F:\\work\\processfile\\processfile\\Debug\\����MES ԭʼ��.xsl.xls"),covOptional ,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional);
    bookMes.AttachDispatch(lpDispMes); 

	lpDisp = books.Open(_T("F:\\work\\processfile\\processfile\\Debug\\all_keywords1207.csv"),covOptional ,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional);
    book.AttachDispatch(lpDisp);


    sheets = book.get_Worksheets(); 
    sheet = sheets.get_Item(COleVariant((short)1));

	sheetsMes = bookMes.get_Worksheets(); 
    sheetMes = sheetsMes.get_Item(COleVariant((short)1));

    //�������Ϊ��A��1���ĵ�Ԫ�� 
    //range = sheet.get_Range(COleVariant(_T("A1")) ,COleVariant(_T("A1")));  

	long rows,rowsMes;
	long temp,tempMes,tempW;
	CString str,strMes,strWrite,strTemp;
     COleVariant vResult;
	 CStdioFile file;
	////��ȡsheet��ʹ�õķ�Χ
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
		file.WriteString("����,�к�,CID,SN\n");
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
			file.WriteString("����,�к�,CID,SN\n");
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
	// TODO: �ڴ���ӿؼ�֪ͨ����������
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
	  sFileMes = _T("����MES ԭʼ��.xsl.xls"); 
	  sFileKeyWord = _T("all_keywords1207.csv");
	  sDriver = GetExcelDriver();

		

     if (sDriver.IsEmpty())
	 {
        // û�з���Excel����
        AfxMessageBox(_T("û�а�װExcel����!"));
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

		 sqlMes = "SELECT SN" //��������˳��    
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
        // ���ݿ���������쳣ʱ...
        AfxMessageBox(_T("���ݿ����: "));
     }
     END_CATCH;
 
#endif

}
