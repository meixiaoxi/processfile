
// processfileDlg.h : ͷ�ļ�
//

#pragma once
#include "CFont0.h"

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
// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
};
