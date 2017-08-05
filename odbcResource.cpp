//odbcResource.cpp
//Copyright (c) 2015 mmYYmmdd

#include "stdafx.h"
#include "odbcResource.hpp"

namespace
{
    bool too_late_to_destruct = false;
}

//****************************************************************
// for DLL_PROCESS_DETACH case of DllMain function
void Too_Late_To_Destruct()
{
    too_late_to_destruct = true;
}
//****************************************************************

namespace mymd  {

odbc_raii_env::odbc_raii_env() noexcept : henv(0)
{   }

odbc_raii_env::~odbc_raii_env() noexcept
{
    if (!too_late_to_destruct)
        ::SQLFreeEnv(henv);
}

bool odbc_raii_env::AllocHandle() noexcept
{
    auto rc = ::SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &henv);
    if (SQL_SUCCESS != rc)  return false;
    rc = ::SQLSetEnvAttr(henv, SQL_ATTR_ODBC_VERSION, reinterpret_cast<void*>(SQL_OV_ODBC3), 0);
    return SQL_SUCCESS ==rc;
}

//********************************************************
odbc_raii_connect::odbc_raii_connect() noexcept : hdbc(0)
{}

odbc_raii_connect::~odbc_raii_connect() noexcept
{
    if (!too_late_to_destruct)
    {
        ::SQLDisconnect(hdbc);
        ::SQLFreeConnect(hdbc);
    }
}

bool odbc_raii_connect::AllocHandle(const odbc_raii_env& env) noexcept
{
    auto const rc = env.invoke(
        [=](HENV x) { return ::SQLAllocHandle(SQL_HANDLE_DBC, x, &hdbc); }
    );
    return SQL_SUCCESS == rc;
}

//********************************************************
odbc_raii_statement::odbc_raii_statement() noexcept : hstmt(0)
{}

odbc_raii_statement::~odbc_raii_statement() noexcept
{
    if (!too_late_to_destruct)
        ::SQLFreeStmt(hstmt, SQL_DROP);
}

tstring
odbc_raii_statement::AllocHandle(const tstring& connectName, const odbc_raii_connect& con)
{
    if (hstmt)  ::SQLFreeStmt(hstmt, SQL_DROP);
    TCHAR ucOutConnectStr[1024];
    SQLSMALLINT ConOut = 0;
    auto pCN = const_cast<SQLTCHAR*>(static_cast<const SQLTCHAR*>(connectName.c_str()));
    auto pCS = static_cast<SQLTCHAR*>(ucOutConnectStr);
    auto len = sizeof(ucOutConnectStr) / sizeof(TCHAR);
    auto pcount = &ConOut;
    auto const expr1 = [=](HDBC x){ return ::SQLDriverConnect(x,
                                                        NULL,
                                                    pCN,
                                                SQL_NTS,
                                            pCS,
                                        static_cast<SQLSMALLINT>(len),
                                    pcount,
                                SQL_DRIVER_NOPROMPT
                            );
                };
    auto const r1 = con.invoke(expr1);
    if ( r1 != SQL_SUCCESS && r1 != SQL_SUCCESS_WITH_INFO )
    {
        SQLDiagRec<SQL_HANDLE_DBC>  diagRec;
        con.invoke(diagRec);
        return diagRec.getMessage();
    }
    ucOutConnectStr[ConOut] = _T('\0');
    auto const r2 = con.invoke(
        [=](HDBC x) { return ::SQLAllocHandle(SQL_HANDLE_STMT, x, &hstmt); }
    );
    if ( r2!=SQL_SUCCESS )
    {
        SQLDiagRec<SQL_HANDLE_DBC>  diagRec;
        con.invoke(diagRec);
        return diagRec.getMessage();
    }
    return tstring(_T(""));
}

//********************************************************
cursor_colser::cursor_colser(const odbc_raii_statement& h, bool b) noexcept : h_(h), close_(b)
{   }

cursor_colser::~cursor_colser() noexcept
{
    if (close_)
        h_.invoke(
            [](HSTMT x) { return ::SQLCloseCursor(x); }
        );
}

//********************************************************

odbc_set::odbc_set(const tstring& connectName, decltype(SQL_CURSOR_FORWARD_ONLY) cursor_type) noexcept
{
    if ( env.AllocHandle() && con.AllocHandle(env) )
        errorMessage_ = st.AllocHandle(connectName, con);
    if ( cursor_type != SQL_CURSOR_FORWARD_ONLY )
    {
        auto ValuePtr = reinterpret_cast<SQLPOINTER>(static_cast<ULONG_PTR>(cursor_type));
        st.invoke(
            [=](HSTMT x) { return ::SQLSetStmtAttr(x, 
                                                   SQL_ATTR_CURSOR_TYPE,
                                                   ValuePtr,
                                                   0);   }
        );
    }
}

odbc_raii_statement& odbc_set::stmt() noexcept
{
    return st;
}

void odbc_set::set_cursor_type(decltype(SQL_CURSOR_STATIC) cursor_type) noexcept
{
    auto ValuePtr = reinterpret_cast<SQLPOINTER>(static_cast<ULONG_PTR>(cursor_type));
    st.invoke(
        [=](HSTMT x) { return ::SQLSetStmtAttr(x, 
                                               SQL_ATTR_CURSOR_TYPE,
                                               ValuePtr,
                                               0);   }
    );
}

bool odbc_set::isError() const noexcept
{
    return 0 < errorMessage_.size();
}

void odbc_set::setErrorMessage(tstring && t) noexcept
{
    errorMessage_ = std::move(t);
}

tstring odbc_set::errorMessage() const
{
    return errorMessage_;
}

//********************************************************

RETCODE execDirect(const tstring& sql_expr, const odbc_raii_statement& stmt) noexcept
{
    auto sql = const_cast<SQLTCHAR*>(static_cast<const SQLTCHAR*>(sql_expr.c_str()));
    return stmt.invoke(
        [=](HSTMT x) { return ::SQLExecDirect(x, sql, SQL_NTS); }
    );
}
//********************************************************

tstring getTypeStr(SQLSMALLINT type) noexcept
{
    tstring ret;
    switch (type)
    {
    case SQL_CHAR:              ret = tstring(_T("CHAR"));          break;
    case SQL_NUMERIC:           ret = tstring(_T("NUMERIC"));       break;
    case SQL_DECIMAL:           ret = tstring(_T("DECIMAL"));       break;
    case SQL_INTEGER:           ret = tstring(_T("INTEGER"));       break;
    case SQL_SMALLINT:          ret = tstring(_T("SMALLINT"));      break;
    case SQL_FLOAT:             ret = tstring(_T("FLOAT"));         break;
    case SQL_REAL:              ret = tstring(_T("REAL"));          break;
    case SQL_DOUBLE:            ret = tstring(_T("DOUBLE"));        break;
    case SQL_VARCHAR:           ret = tstring(_T("VARCHAR"));       break;
    case SQL_TYPE_DATE:         ret = tstring(_T("TYPE_DATE"));     break;
    case SQL_TYPE_TIME:         ret = tstring(_T("TYPE_TIME"));     break;
    case SQL_TYPE_TIMESTAMP:    ret = tstring(_T("TYPE_TIMESTAMP")); break;
    case SQL_WLONGVARCHAR:      ret = tstring(_T("WLONGVARCHAR"));  break;
    case SQL_WVARCHAR:          ret = tstring(_T("WVARCHAR"));      break;
    case SQL_WCHAR:             ret = tstring(_T("WCHAR"));         break;
    case SQL_BIT:               ret = tstring(_T("BIT"));           break;
    case SQL_TINYINT:           ret = tstring(_T("TINYINT"));       break;
    case SQL_BIGINT:            ret = tstring(_T("BIGINT"));        break;
    case SQL_LONGVARBINARY:     ret = tstring(_T("LONGVARBINARY")); break;
    case SQL_VARBINARY:         ret = tstring(_T("VARBINARY"));     break;
    case SQL_BINARY:            ret = tstring(_T("BINARY"));        break;
    case SQL_LONGVARCHAR:       ret = tstring(_T("LONGVARCHAR"));   break;
    default:                    ret = tstring(_T("?"));
    }
    return ret;
}

}   // namespace mymd
