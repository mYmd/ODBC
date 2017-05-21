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

namespace mymd  {

using tstring = std::basic_string<TCHAR>;

//**************************************************************
class odbc_raii_env {
    HENV    henv;
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
    {   return (std::forward<T>(expr))(henv);   }
};

//**************************************************************
class odbc_raii_connect {
    HDBC    hdbc;
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
    {   return (std::forward<T>(expr))(hdbc);   }
};

//**************************************************************
class odbc_raii_statement   {
    HSTMT   hstmt;
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
    {   return (std::forward<T>(expr))(hstmt);  }
};

//**************************************************************
class cursor_colser {
    const odbc_raii_statement&  h_;
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
    tstring             errorMessage_;
public:
    odbc_set(const tstring& connectName);
    odbc_raii_statement&  stmt();
    bool isError() const;
    void setErrorMessage(tstring&&);
    tstring errorMessage() const;
};

//**************************************************************
tstring getTypeStr(SQLSMALLINT);
//********************************************************

struct column_t {
    static std::size_t const nameSize = 256;
    using name_type = std::array<TCHAR, nameSize>;
    using buffer_type = std::basic_string<UCHAR>;
    static std::size_t const bufferSize = 16384;
};
//********************************************************

//Diagnostic Message 診断メッセージ
template <SQLSMALLINT HandleType = SQL_HANDLE_STMT, std::size_t bufferSize = 1024>
class SQLDiagRec   {
    SQLSMALLINT recNum;
    SQLWCHAR SQLState[6];
    SQLWCHAR szErrorMsg[bufferSize];
public:
    SQLDiagRec() : recNum{1}
    {   SQLState[0] = _T('\0'); szErrorMsg[0] = _T('\0');   }
    void setnum(SQLSMALLINT a)      {   recNum = a;     }
    tstring getMessage() const      {   return szErrorMsg;  }
    tstring getState() const        {   return SQLState;    }
    RETCODE operator ()(HSTMT x)
    {
        SQLSMALLINT o_o;
        return ::SQLGetDiagRec(HandleType, x, recNum, SQLState, NULL, szErrorMsg, bufferSize, &o_o);
    }
};

//********************************************************

// カタログ関数
template <typename FC, typename FP>
std::size_t catalogValue(
    FC&&                        catalog_func    ,   //
    odbc_raii_statement const&  st              ,
    SQLUSMALLINT                ColumnNumber    ,
    FP&&                        push_back_func  )   // <- TCHAR const* p
{
    auto result = st.invoke(std::forward<FC>(catalog_func));
    if (SQL_SUCCESS != result)      return 0;
    SQLSMALLINT nresultcols{0};
    {
        SQLSMALLINT* pl = &nresultcols;
        auto result = st.invoke(
            [=](HSTMT x) { return ::SQLNumResultCols(x, pl); }
        );
    }
    const std::size_t ColumnNameLen = column_t::nameSize;
    SQLCHAR  rgbValue[ColumnNameLen];
    SQLLEN   pcbValue{0};
    cursor_colser   c_closer(st, true);
    {
        auto p_rgbValue = static_cast<SQLPOINTER>(rgbValue);
        auto p_pcbValue = &pcbValue;
        for ( auto j = 0; j < nresultcols; ++j )
        {
            if (j == ColumnNumber)    continue;
            auto result = st.invoke(
                [=](HSTMT x) { return ::SQLBindCol(x, j, SQL_C_CHAR, NULL, 0, NULL); }
            );
        }
        auto result = st.invoke(
            [=](HSTMT x) { return ::SQLBindCol( x,
                                            ColumnNumber,      //COLUMN_NAME
                                        SQL_C_CHAR,
                                    p_rgbValue,
                                ColumnNameLen,
                            p_pcbValue);
            }
        );
        if (SQL_SUCCESS != result)      return 0;
    }
    auto SQLFetch_expr = [=](HSTMT x) { return ::SQLFetch(x); };
    std::size_t counter{0};
    TCHAR tcharBuffer[ColumnNameLen];
    while (true)
    {
        tcharBuffer[0] = _T('\0');
        auto fetch_result = st.invoke(SQLFetch_expr);
        if ((SQL_SUCCESS != fetch_result) && (SQL_SUCCESS_WITH_INFO != fetch_result))
            break;
        int mb = ::MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, (LPCSTR)rgbValue, -1, tcharBuffer, ColumnNameLen);
        std::forward<FP>(push_back_func)((1 < mb) ? tcharBuffer : nullptr);
        ++counter;
    }
    return counter;
}

//******************************************************************

//  SELECT  ,  INSERT  ,  UPDATE  ,  ...
RETCODE execDirect(const tstring& sql_expr, const odbc_raii_statement& stmt);

//******************************************************************
template <typename F>
SQLSMALLINT columnAttribute(odbc_raii_statement const&          stmt    ,
                        tstring const&                      sql_expr,
                    std::vector<column_t::buffer_type>* pBuffer ,
                std::vector<SQLLEN>*                pdatastrlen,
            F&&                                 write_func,
        bool close_)
{
    auto const rc = execDirect(sql_expr, stmt);
    if (rc == SQL_ERROR || rc == SQL_INVALID_HANDLE)    return 0;
    SQLSMALLINT nresultcols{0};
    {
        SQLSMALLINT* pl = &nresultcols;
        RETCODE const rc = stmt.invoke(
            [=](HSTMT x) { return ::SQLNumResultCols(x, pl); }
        );
        if (SQL_SUCCESS != rc)  return 0;
    }
    std::vector<column_t::name_type>    colname(nresultcols);
    std::vector<SQLSMALLINT>            colnamelen(nresultcols);
    std::vector<SQLULEN>                collen(nresultcols);
    std::vector<SQLSMALLINT>            nullable(nresultcols);
    std::vector<SQLSMALLINT>            coltype(nresultcols);
    std::vector<SQLSMALLINT>            scale(nresultcols);
    if (pBuffer)
    {
        pBuffer->clear();
        pBuffer->resize(nresultcols);
    }
    if (pdatastrlen)
    {
        pdatastrlen->clear();
        pdatastrlen->resize(nresultcols);
    }
    int j = 0;
    auto SQLDescribeColExpr = [&](HSTMT x)    {
        return ::SQLDescribeCol(x                           ,
                                static_cast<UWORD>(j+1)     ,
                                colname[j].data()           ,
                                static_cast<SQLSMALLINT>(column_t::nameSize * sizeof(TCHAR)),
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
                            &(*pdatastrlen)[j]);
    };
    const std::size_t StrSizeofColumn = column_t::bufferSize;
    cursor_colser   c_closer(stmt, close_);
    for ( j = 0; j < nresultcols; ++j )
    {
        RETCODE rc = stmt.invoke(SQLDescribeColExpr);
        if (pBuffer && pdatastrlen)
        {
            auto dlen = collen[j];
            (*pBuffer)[j].resize((0 < dlen && dlen < StrSizeofColumn) ? dlen+1 : StrSizeofColumn);
            rc = stmt.invoke(SQLBindColExpr);
        }
    }
    std::forward<F>(write_func)(colname, colnamelen, collen, nullable, coltype, scale);
    return nresultcols;
}

//******************************************************************

    struct no_header {
        void operator()(
            std::vector<column_t::name_type>&,
            std::vector<SQLSMALLINT>&,
            std::vector<SQLULEN>&,
            std::vector<SQLSMALLINT>&,
            std::vector<SQLSMALLINT>&,
            std::vector<SQLSMALLINT>&)  {  }
    };

    struct bool_sentinel    {
        explicit operator bool() const  { return true; }
        friend bool operator ,(bool b, const bool_sentinel&)    { return b; }
    };

//******************************************************************

template <typename FH, typename FI, typename FE, typename FA>
std::size_t select_table(   odbc_raii_statement const& stmt , 
                            tstring const&  sql_expr        , 
                            FH&&            header_func     ,
                            FI&&            init_func       , 
                            FE&&            elem_func       ,
                            FA&&            add_func        )
{
    std::vector<SQLSMALLINT>        coltype;
    auto write_func = [&] ( std::vector<column_t::name_type>&   colname_    ,
                            std::vector<SQLSMALLINT>&           colnamelen_ ,
                            std::vector<SQLULEN>&               collen_     ,
                            std::vector<SQLSMALLINT>&           nullable_   ,
                            std::vector<SQLSMALLINT>&           coltype_    ,
                            std::vector<SQLSMALLINT>&           scale_      )
    {
        coltype = coltype_;
        std::forward<FH>(header_func)(colname_, colnamelen_, collen_, nullable_, coltype_, scale_);
    };
    std::vector<column_t::buffer_type>  buffer;
    std::vector<SQLLEN>                 datastrlen;
    SQLSMALLINT nresultcols = columnAttribute(  stmt        ,
                                            sql_expr    ,
                                        &buffer     ,
                                    &datastrlen ,
                                write_func  ,
                            false       );
    if (nresultcols == 0 )          return 0;
    //-----------------------------------------------
    cursor_colser   c_closer(stmt, true);
    bool_sentinel   bp;
    if ( !(std::forward<FI>(init_func)(nresultcols), bp) )
        return 0;
    const std::size_t StrSizeofColumn = column_t::bufferSize;
    TCHAR tcharBuffer[StrSizeofColumn];
    auto const fetch_expr = [](HSTMT x) { return ::SQLFetch(x); };
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
                int mb = ::MultiByteToWideChar( CP_ACP,
                                            MB_PRECOMPOSED,
                                        reinterpret_cast<LPCSTR>(&buffer[j][0]),
                                    -1,
                                tcharBuffer,
                            StrSizeofColumn);
                std::forward<FE>(elem_func)(j, (SQL_NULL_DATA == datastrlen[j]) ? nullptr: tcharBuffer, coltype[j]);
            }
            if ( !( std::forward<FA>(add_func)(counter++), bp) )
                break;
        }
        else
        {
            break;
        }
    }
    return counter;
}

}   // namespace mymd
