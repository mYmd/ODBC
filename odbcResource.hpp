//odbcResource.hpp
//Copyright (c) 2015 mmYYmmdd

#pragma once

#include <sql.h>
#include <sqlext.h>
#include <odbcinst.h>
#include <tchar.h>
#include <string>
#include <vector>
#include <array>

#pragma comment(lib, "odbccp32.lib")

using tstring = std::basic_string<TCHAR>;

//**************************************************************
class odbc_raii_env	{
	HENV	henv;
	odbc_raii_env(const odbc_raii_env&) = delete;
	odbc_raii_env(odbc_raii_env&&) = delete;
	odbc_raii_env& operator =(const odbc_raii_env&) = delete;
	odbc_raii_env& operator =(odbc_raii_env&&) = delete;
public:
	odbc_raii_env();
	~odbc_raii_env();
	void AllocHandle();
	template <typename T>
	RETCODE invoke(T&& expr) const
	{	return (std::forward<T>(expr))(henv);	}
};

//**************************************************************
class odbc_raii_connect	{
	HDBC	hdbc;
	odbc_raii_connect(const odbc_raii_connect&) = delete;
	odbc_raii_connect(odbc_raii_connect&&) = delete;
	odbc_raii_connect& operator =(const odbc_raii_connect&) = delete;
	odbc_raii_connect& operator =(odbc_raii_connect&&) = delete;
public:
	odbc_raii_connect();
	~odbc_raii_connect();
	void AllocHandle(const odbc_raii_env& env);
	template <typename T>
	RETCODE invoke(T&& expr) const
	{	return (std::forward<T>(expr))(hdbc);	}
};

//**************************************************************
class odbc_raii_statement	{
	HSTMT	hstmt;
	odbc_raii_statement(const odbc_raii_statement&) = delete;
	odbc_raii_statement(odbc_raii_statement&&) = delete;
	odbc_raii_statement& operator =(const odbc_raii_statement&) = delete;
	odbc_raii_statement& operator =(odbc_raii_statement&&) = delete;
public:
	odbc_raii_statement();
	~odbc_raii_statement();
	tstring AllocHandle(const tstring& connectName, const odbc_raii_connect& con);
	template <typename T>
	RETCODE invoke(T&& expr) const
	{	return (std::forward<T>(expr))(hstmt);	}
};

//**************************************************************
class cursor_colser	{
	const odbc_raii_statement&	h_;
public:
	cursor_colser(const odbc_raii_statement& h);
	~cursor_colser();
};

//**************************************************************

class odbc_set {
    odbc_raii_env       env;
    odbc_raii_connect   con;
    odbc_raii_statement st;
    __int32*            pNo;
public:
    odbc_set(const tstring& connectName, __int32& myNo);
    ~odbc_set();
    odbc_raii_statement&  stmt();
};

//**************************************************************
class odbc_raii_select	{
	odbc_raii_select(const odbc_raii_select&) = delete;
	odbc_raii_select(odbc_raii_select&&) = delete;
	odbc_raii_select& operator =(const odbc_raii_select&) = delete;
	odbc_raii_select& operator =(odbc_raii_select&&) = delete;
public:
	using result_type = std::vector<std::vector<tstring>>;
    static const std::size_t    StrSizeofColumn = 16384;
    static const std::size_t    ColumnNameLen = 256;
    using column_name_type = std::array<TCHAR, ColumnNameLen>;
    odbc_raii_select()  {}
    ~odbc_raii_select() {}
    RETCODE execDirect(const tstring&, const odbc_raii_statement&) const;
    SQLSMALLINT columnAttribute(const tstring&                  ,
                                const odbc_raii_statement&      ,
                                std::vector<column_name_type>&  ,
                                std::vector<SQLSMALLINT>&       ,
                                std::vector<SQLULEN>&           ,
                                std::vector<SQLSMALLINT>&       ,
                                std::vector<SQLSMALLINT>&       ,
                                std::vector<SQLSMALLINT>&       ,
                                std::vector<SQLLEN>&            ,
                                std::vector<std::basic_string<UCHAR>>*       
                                ) const;
    result_type select( int                             timeOutSec,
                        const tstring&                  sql_expr,
                        const odbc_raii_statement&      stmt,
                        std::vector<column_name_type>*  pColname,
                        std::vector<SQLSMALLINT>*       pColnamelen,
                        std::vector<SQLULEN>*           pCollen,
                        std::vector<SQLSMALLINT>*       pNullable,
                        std::vector<SQLSMALLINT>*       pColtype,
                        std::vector<SQLSMALLINT>*       pScale,
                        std::vector<SQLLEN>*            pDatastrlen
                        ) const;
};

