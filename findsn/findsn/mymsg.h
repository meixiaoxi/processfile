#ifndef __MY_MSG_H
#define __MY_MSG_H

typedef struct LogBuf
{
	char *buf;
	int len;
}LOGBUF, *pLOGBUF;



#define WM_USER_NOTIFY			(WM_USER+100)



#define WP_START_TEST					1
#define WP_PRODUCT_IN_WARRANTY			2		// 产品在保修期内
#define WP_PRODUCT_EXPIRED_WARRANTY		3	// 产品超出保修期
#define WP_FAIL_OPEN_SN_SOURCE_FILE		4
#define WP_SAVE_INFO_FILE_PASS			5
#define WP_SAVE_INFO_FILE_FAIL			6
#define WP_PRINT_LOG_STR				7
#define WP_START_EXCEL_APP				8				
#endif