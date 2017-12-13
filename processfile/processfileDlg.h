
// processfileDlg.h : 头文件
//

#pragma once
#include "CFont0.h"

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
    CFont0 font;   
    CRange cols; 
    LPDISPATCH lpDisp;

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

	BOOL CprocessfileDlg::mOperate();
	//CString GetCellString(long irow, long icolumn);
// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
};
