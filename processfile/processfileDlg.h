
// processfileDlg.h : 头文件
//

#pragma once
#include "CFont0.h"
#include<odbcinst.h> 
#include<afxdb.h>



struct threadInfo
{
	HWND	hWnd;//主窗口句柄，用于消息的发送
	int		index;//线程标号
	int		portnum;//串口号
	CWinThread *pThread;
	int		status;
	CRange RangeMes;
	CRange RangeRow;
	CRange RangeFind;
	CRange tRange;
	long cnt;
};


// CprocessfileDlg 对话框
class CprocessfileDlg : public CDialogEx
{
// 构造
public:
	CprocessfileDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_PROCESSFILE_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持
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
// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnCancel();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
};

