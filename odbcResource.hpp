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

//*******************************************************
tstring getTypeStr(SQLSMALLINT);

// ÉJÉ^ÉçÉOä÷êî
template <typename FC, typename FP>
std::size_t catalogValue(
    FC&&                        catalog_func,       //
    odbc_raii_statement const&  st,
    SQLUSMALLINT                ColumnNumber,
    FP&&                        push_back_func      // <- TCHAR const* p
)
{
    cursor_colser   c_closer(st);
    auto result = st.invoke(std::forward<FC>(catalog_func));
    if (SQL_SUCCESS != result)
        return 0;
    SQLSMALLINT nresultcols = 0;
    {
        SQLSMALLINT* pl = &nresultcols;
        auto result = st.invoke(
            [=](HSTMT x) { return ::SQLNumResultCols(x, pl); }
        );
    }
    SQLCHAR  rgbValue[odbc_raii_select::ColumnNameLen];
    SQLLEN   pcbValue;
    {
        auto p_rgbValue = static_cast<SQLPOINTER>(rgbValue);
        auto p_pcbValue = &pcbValue;
        for (auto j = 0; j < nresultcols; ++j)
        {
            if (j == ColumnNumber)    continue;
            auto result = st.invoke(
                [=](HSTMT x) { return ::SQLBindCol(x, j, SQL_C_CHAR, NULL, 0, NULL); }
            );
        }
        auto result = st.invoke(
            [=](HSTMT x) { return ::SQLBindCol(x,
                ColumnNumber,      //COLUMN_NAME
                SQL_C_CHAR,
                p_rgbValue,
                odbc_raii_select::ColumnNameLen,
                p_pcbValue);
        }
        );
        if (SQL_SUCCESS != result)
            return 0;
    }
    auto SQLFetch_expr = [=](HSTMT x) { return ::SQLFetch(x); };
    std::vector<VARIANT> vec;
    std::size_t counter{0};
    while (true)
    {
        auto fetch_result = st.invoke(SQLFetch_expr);
        if ((SQL_SUCCESS != fetch_result) && (SQL_SUCCESS_WITH_INFO != fetch_result))
            break;
        TCHAR tcharBuffer[odbc_raii_select::ColumnNameLen];
        int mb = ::MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, (LPCSTR)rgbValue, -1, tcharBuffer, odbc_raii_select::ColumnNameLen);
        tstring const str(tcharBuffer);
        TCHAR const* p = str.empty() ? 0 : &str[0];
        std::forward<FP>(push_back_func)(p);
        ++counter;
    }
    return counter;
}
