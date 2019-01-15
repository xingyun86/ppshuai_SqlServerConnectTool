// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//  are changed infrequently
//

#pragma once

#ifndef STRICT
#define STRICT
#endif

#include "targetver.h"

#define _ATL_APARTMENT_THREADED

#include <atlbase.h>
#include <atlapp.h>

extern CServerAppModule _Module;

// This is here only to tell VC7 Class Wizard this is an ATL project
#ifdef ___VC7_CLWIZ_ONLY___
CComModule
CExeModule
#endif

#include <atlcom.h>
#include <atlhost.h>
#include <atlwin.h>
#include <atlctl.h>

#include <atlframe.h>
#include <atlctrls.h>
#include <atldlgs.h>


#if defined(_WIN64) || defined(WIN64)
#import "C:/Program Files/Common Files/System/ado/msado15.dll" no_namespace rename("EOF","adoEOF")rename("BOF","adoBOF")
#elif defined(_WIN32) || defined(WIN32)
#import "C:/Program Files (x86)/Common Files/System/ado/msado15.dll" no_namespace rename("EOF","adoEOF")rename("BOF","adoBOF")
#else
#error Undefined operatoration system bits 32 or 64.
#endif

class CAdoSql
{
public:
	CAdoSql()
	{
		m_pConnectionPtr = NULL;
		m_pRecordSetPtr = NULL;
		m_pCommandPtr = NULL;
		//创建Connection对象
		//m_hResult = m_pConnection.CreateInstance("ADODB.Connection");
		m_hResult = m_pConnectionPtr.CreateInstance(__uuidof(Connection));
		if (SUCCEEDED(m_hResult))
		{
			//创建记录集指针对象实例
			//m_hResult = m_pRecordSet.CreateInstance("ADODB.Recordset");
			m_hResult = m_pRecordSetPtr.CreateInstance(__uuidof(Recordset));
			if (SUCCEEDED(m_hResult))
			{
				//m_hResult = m_pCommandPtr.CreateInstance("ADODB.Command");
				m_hResult = m_pCommandPtr.CreateInstance(__uuidof(Command));
			}
		}
	}
	virtual ~CAdoSql()
	{
	}
	HRESULT CoInit()
	{
		//初始化COM组件
		return ::CoInitialize(NULL);
	}
	HRESULT CoInitMultiThread()
	{
		//初始化COM组件
		return ::CoInitializeEx(NULL, COINIT_MULTITHREADED);
	}
	HRESULT CoExit()
	{
		//析构释放COM组件
		::CoUninitialize();
		return S_OK;
	}
	//连接数据库
	HRESULT InitConn(LPTSTR lpszStrConn, LPCTSTR lpUserId = NULL, LPCTSTR lpPassword = NULL, LONG nOptions = adModeUnknown, LONG nConnectionTimeout = (-1), LONG nCommandTimeout = (-1))
	{
		//1、新建一个文件，名字任意取，后缀名必须为udl（如何hello.udl）
		//2、双击hello.udl文件，进入数据连接属性面板，填写好数据源（选择自己创建的数据源名字的）
		//3、再以记事本方式打开，第三行的数据就是连接字符
		USES_CONVERSION_EX;
		m_pConnectionPtr->ConnectionTimeout = nConnectionTimeout < 0 ? m_pConnectionPtr->GetConnectionTimeout() : nConnectionTimeout;
		m_pConnectionPtr->CommandTimeout = nCommandTimeout < 0 ? m_pConnectionPtr->GetCommandTimeout() : nCommandTimeout;
		return m_pConnectionPtr->Open(T2W_EX((LPTSTR)lpszStrConn, _tcslen(lpszStrConn)), lpUserId ? T2W_EX((LPTSTR)lpUserId, _tcslen(lpUserId)) : lpUserId, lpPassword ? T2W_EX((LPTSTR)lpPassword, _tcslen(lpPassword)) : lpPassword, nOptions);
	}
	//断开数据库
	HRESULT ExitConn()
	{
		HRESULT hResult = S_OK;
		//关闭记录集
		if (m_pRecordSetPtr->GetState() != adStateClosed)
		{
			hResult = m_pRecordSetPtr->Close();
			m_pRecordSetPtr = NULL;
		}
		if (SUCCEEDED(hResult))
		{
			if (m_pConnectionPtr->GetState() == adStateOpen)
			{
				//关闭连接
				hResult = m_pConnectionPtr->Close();
			}
		}
		return hResult;
	}
	//获得记录集
	HRESULT GetRecordSets(LPCTSTR lpszSqlCmd, CursorTypeEnum nCursorType = adOpenDynamic, LockTypeEnum nLockType = adLockOptimistic, LONG nOptions = adCmdText)
	{
		//打开记录集
		USES_CONVERSION_EX;
		return m_pRecordSetPtr->Open(T2W_EX((LPTSTR)lpszSqlCmd, _tcslen(lpszSqlCmd)), m_pConnectionPtr.GetInterfacePtr(), nCursorType, nLockType, nOptions);
	}
	//执行SQL存储过程
	HRESULT ExcuteCommand(LPCTSTR lpszSqlCmd, LONG nCommandTimeout = (-1), LONG nOptions = adCmdText)
	{
		USES_CONVERSION_EX;
		m_pCommandPtr->ActiveConnection = m_pConnectionPtr;
		m_pCommandPtr->CommandText = lpszSqlCmd;
		m_pCommandPtr->CommandTimeout = nCommandTimeout < 0 ? m_pCommandPtr->GetCommandTimeout() : nCommandTimeout;;
		_variant_t RecordsAffected;
		m_pRecordSetPtr = m_pCommandPtr->Execute(NULL, NULL, nOptions);
		return m_pRecordSetPtr ? S_OK : S_FALSE;
	}

	//执行SQL语句
	HRESULT ExcuteCommand(LPCTSTR lpszSqlCmd, LONG nOptions = adCmdText)
	{
		USES_CONVERSION_EX;
		_variant_t RecordsAffected;
		m_pRecordSetPtr = m_pConnectionPtr->Execute(T2W_EX((LPTSTR)lpszSqlCmd, _tcslen(lpszSqlCmd)), &RecordsAffected, nOptions);
		return m_pRecordSetPtr ? S_OK : S_FALSE;
	}

public:
	HRESULT m_hResult;
	_ConnectionPtr m_pConnectionPtr;  //连接数据库对象指针
	_RecordsetPtr  m_pRecordSetPtr;   //数据集对象指针
	_CommandPtr  m_pCommandPtr;   //命令对象指针
};

#ifdef _UNICODE
#if defined _M_IX86
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif
#endif

