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
    bool close_;
public:
	cursor_colser(const odbc_raii_statement& h, bool b);
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
tstring getTypeStr(SQLSMALLINT);
//********************************************************
using buffer_t = std::basic_string<UCHAR>;
using column_name_type = std::array<TCHAR, 256>;
using result_type = std::vector<std::vector<tstring>>;
//********************************************************


// ÉJÉ^ÉçÉOä÷êî
template <typename FC, typename FP>
std::size_t catalogValue(
    FC&&                        catalog_func    ,   //
    odbc_raii_statement const&  st              ,
    SQLUSMALLINT                ColumnNumber    ,
    FP&&                        push_back_func  )   // <- TCHAR const* p
{
    cursor_colser   c_closer(st, true);
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
    const std::size_t ColumnNameLen = 256;
    SQLCHAR  rgbValue[ColumnNameLen];
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
                ColumnNameLen,
                p_pcbValue);
        }
        );
        if (SQL_SUCCESS != result)
            return 0;
    }
    auto SQLFetch_expr = [=](HSTMT x) { return ::SQLFetch(x); };
    std::vector<VARIANT> vec;
    std::size_t counter{0};
    TCHAR tcharBuffer[ColumnNameLen];
    while (true)
    {
        tcharBuffer[0] = _T('\0');
        auto fetch_result = st.invoke(SQLFetch_expr);
        if ((SQL_SUCCESS != fetch_result) && (SQL_SUCCESS_WITH_INFO != fetch_result))
            break;
        int mb = ::MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, (LPCSTR)rgbValue, -1, tcharBuffer, ColumnNameLen);
        tstring str(tcharBuffer);
        TCHAR const* p = mb ? &str[0]: nullptr;
        std::forward<FP>(push_back_func)(p);
        ++counter;
    }
    return counter;
}

RETCODE execDirect(const tstring& sql_expr, const odbc_raii_statement& stmt);

template <typename F>
SQLSMALLINT columnAttribute(odbc_raii_statement const&  stmt    ,
                        tstring const&              sql_expr,
                    std::vector<buffer_t>*      pBuffer ,
                F&&                         write_func,
            bool close_)
{
    const std::size_t ColumnNameLen = 256;
    {
        auto const rc = execDirect(sql_expr, stmt);
        if (rc == SQL_ERROR || rc == SQL_INVALID_HANDLE)
            return 0;
    }
    SQLSMALLINT nresultcols = 0;
    {
        SQLSMALLINT* pl = &nresultcols;
        RETCODE const rc = stmt.invoke(
            [=](HSTMT x) { return ::SQLNumResultCols(x, pl); }
        );
        if (SQL_SUCCESS != rc)  return 0;
    }
    std::vector<column_name_type>   colname(nresultcols);
    std::vector<SQLSMALLINT>        colnamelen(nresultcols);
    std::vector<SQLULEN>            collen(nresultcols);
    std::vector<SQLSMALLINT>        nullable(nresultcols);
    std::vector<SQLSMALLINT>        coltype(nresultcols);
    std::vector<SQLSMALLINT>        scale(nresultcols);
    std::vector<SQLLEN>             datastrlen(nresultcols);
    if (pBuffer)
    {
        pBuffer->clear();
        pBuffer->resize(nresultcols);
    }
    int j = 0;
    auto SQLDescribeColExpr = [&](HSTMT x)    {
        return ::SQLDescribeCol(x                           ,
                                static_cast<UWORD>(j+1)     ,
                                colname[j].data()           ,
                                ColumnNameLen*sizeof(TCHAR) ,
                                &colnamelen[j]              ,
                                &coltype[j]                 ,
                                &collen[j]                  ,
                                &scale[j]                   ,
                                &nullable[j]                );
    };
    auto SQLBindColExpr = [&](HSTMT x) {
        return ::SQLBindCol(x, 
                            static_cast<UWORD>(j+1),
                            SQL_C_CHAR, 
                            &(*pBuffer)[j][0],
                            (*pBuffer)[j].size() * sizeof(UCHAR),
                            &datastrlen[j]);
    };
    const std::size_t StrSizeofColumn = 16384;
    cursor_colser   c_closer(stmt, close_);
    for ( j = 0; j < nresultcols; ++j )
    {
        RETCODE rc = stmt.invoke(SQLDescribeColExpr);
        if (pBuffer)
        {
            auto dlen = collen[j];
            (*pBuffer)[j].resize((0 < dlen && dlen < StrSizeofColumn) ? dlen+1 : StrSizeofColumn);
            rc = stmt.invoke(SQLBindColExpr);
        }
    }
    std::forward<F>(write_func)(colname, colnamelen, collen, nullable, coltype, scale, datastrlen);
    return nresultcols;
}

template <typename FH, typename FI, typename FE, typename FA>
std::size_t select_table(   odbc_raii_statement const& stmt , 
                            tstring const&  sql_expr        , 
                            FH&&            header_func     ,
                            FI&&            init_func       , 
                            FE&&            elem_func       ,
                            FA&&            add_func        )
{
    std::vector<SQLSMALLINT>        coltype;
    std::vector<SQLLEN>             datastrlen;
    auto write_func = [&] ( std::vector<column_name_type>&  colname_    ,
                            std::vector<SQLSMALLINT>&       colnamelen_ ,
                            std::vector<SQLULEN>&           collen_     ,
                            std::vector<SQLSMALLINT>&       nullable_   ,
                            std::vector<SQLSMALLINT>&       coltype_    ,
                            std::vector<SQLSMALLINT>&       scale_      ,
                            std::vector<SQLLEN>&            datastrlen_)
    {
        std::forward<FH>(header_func)(colname_, colnamelen_, collen_, nullable_, coltype_, scale_, datastrlen_);
        coltype     = std::move(coltype_);
        datastrlen  = std::move(datastrlen_);
    };
    cursor_colser   c_closer(stmt, true);
    std::vector<buffer_t> buffer;
    SQLSMALLINT nresultcols = columnAttribute(  stmt        ,
                                            sql_expr    ,
                                        &buffer     , 
                                    write_func  ,
                                false       );
    if (nresultcols == 0 )          return 0;
    //-----------------------------------------------
    std::forward<FI>(init_func)(nresultcols);
    const std::size_t StrSizeofColumn = 16384;
    TCHAR tcharBuffer[StrSizeofColumn];
    std::vector<tstring>    record(nresultcols);
    auto const fetch_expr = [](HSTMT x) { return SQLFetch(x); };
    std::size_t counter{0};
    while (true)
    {
        for (int j = 0; j < nresultcols; ++j)
            buffer[j][0] = '\0';
        RETCODE const rc = stmt.invoke(fetch_expr);
        if (rc == SQL_SUCCESS || rc == SQL_SUCCESS_WITH_INFO)
        {
            for (SQLSMALLINT j = 0; j < nresultcols; ++j)
            {
                if (0 < datastrlen[j] && datastrlen[j] < long(buffer[j].size()))
                    buffer[j][datastrlen[j]] = '\0';
                int mb = ::MultiByteToWideChar( CP_ACP,
                                            MB_PRECOMPOSED,
                                        (LPCSTR)&buffer[j][0],
                                    -1,
                                tcharBuffer,
                            StrSizeofColumn);
                record[j] = tcharBuffer;
                std::forward<FE>(elem_func)(j, tstring{tcharBuffer}, coltype[j]);
            }
            std::forward<FA>(add_func)(counter++);
        }
        else
        {
            break;
        }
    }
    return counter;
}
