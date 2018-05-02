
// processfileDlg.h : ͷ�ļ�
//

#pragma once
#include "CFont0.h"
#include<odbcinst.h> 
#include<afxdb.h>



struct threadInfo
{
	HWND	hWnd;//�����ھ����������Ϣ�ķ���
	int		index;//�̱߳��
	int		portnum;//���ں�
	CWinThread *pThread;
	int		status;
	CRange RangeMes;
	CRange RangeRow;
	CRange RangeFind;
	CRange tRange;
	long cnt;
};


// CprocessfileDlg �Ի���
class CprocessfileDlg : public CDialogEx
{
// ����
public:
	CprocessfileDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_PROCESSFILE_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��
	CApplication app; 
    CWorkbook book;  
    CWorkbooks books;  
    CWorksheet sheet;  
    CWorksheets sheets; 
    CRange range; 
	CRange tRange;
	CRange rangeFind;
	CRange rangeRow;
    CFont0 font;   
    CRange cols; 
    LPDISPATCH lpDisp;
	LPDISPATCH lpDispFind;

	//MES
	CApplication appMes; 
    CWorkbook bookMes;  
    CWorkbooks booksMes;  
    CWorksheet sheetMes;  
    CWorksheets sheetsMes; 
    CRange rangeMes; 
	CRange tRangeMes;
    CFont0 fontMes;   
    CRange colsMes; 
    LPDISPATCH lpDispMes;

	char mDesFile[200];
	char mFileMes[200];
	char mFile[200];

	BOOL CprocessfileDlg::mOperate();
	void CprocessfileDlg::mFindCidSN();
	threadInfo m_info[50];
	//CString GetCellString(long irow, long icolumn);
// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnCancel();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
};

