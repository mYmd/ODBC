//odbcResource.cpp
//Copyright (c) 2015 mmYYmmdd

#include "stdafx.h"
#include "odbcResource.hpp"

namespace
{
    RETCODE do_SQLDescribeCol(  const odbc_raii_statement&,
                                UWORD           ,
                                TCHAR*          ,
                                SWORD           ,
                                SQLSMALLINT*    ,
                                SQLSMALLINT*    ,
                                SQLULEN*        ,
                                SQLSMALLINT*    ,
                                SQLSMALLINT*    );

    RETCODE do_SQLBindCol(  const odbc_raii_statement&  ,
                            UWORD                       ,
                            UCHAR*                      ,
                            std::size_t                 ,
                            SQLLEN*                     );

    bool too_late_to_destruct = false;
}

//****************************************************************
// for DLL_PROCESS_DETACH case of DllMain function
void Too_Late_To_Destruct()
{
    too_late_to_destruct = true;
}
//****************************************************************

odbc_raii_env::odbc_raii_env() : henv(0)
{   }

odbc_raii_env::~odbc_raii_env()
{
    if (!too_late_to_destruct)
        SQLFreeEnv(henv);
}

void odbc_raii_env::AllocHandle()
{
    RETCODE rc = ::SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &henv);
    if (SQL_SUCCESS != rc) throw rc;
    rc = ::SQLSetEnvAttr(henv, SQL_ATTR_ODBC_VERSION, reinterpret_cast<void*>(SQL_OV_ODBC3), 0);
    if (SQL_SUCCESS !=rc) throw rc;
}

//********************************************************
odbc_raii_connect::odbc_raii_connect() : hdbc(0)
{}

odbc_raii_connect::~odbc_raii_connect()
{
    if (!too_late_to_destruct)
    {
        ::SQLDisconnect(hdbc);
        ::SQLFreeConnect(hdbc);
    }
}

void odbc_raii_connect::AllocHandle(const odbc_raii_env& env)
{
    RETCODE const rc = env.invoke(
        [=](HENV x) { return ::SQLAllocHandle(SQL_HANDLE_DBC, x, &hdbc); }
    );
    if (SQL_SUCCESS != rc) throw rc;
}

//********************************************************
odbc_raii_statement::odbc_raii_statement() : hstmt(0)
{}

odbc_raii_statement::~odbc_raii_statement()
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
    SQLTCHAR*   pCN = const_cast<SQLTCHAR*>(static_cast<const SQLTCHAR*>(connectName.c_str()));
    SQLTCHAR*   pCS = static_cast<SQLTCHAR*>(ucOutConnectStr);
    std::size_t const len = sizeof(ucOutConnectStr) / sizeof(TCHAR);
    SQLSMALLINT* pcount = &ConOut;
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
    ucOutConnectStr[ConOut] = _T('\0');
    auto const r2 = con.invoke(
        [=](HDBC x) { return ::SQLAllocHandle(SQL_HANDLE_STMT, x, &hstmt); }
    );
    if ( r2!=SQL_SUCCESS )  throw r2;
    return tstring(ucOutConnectStr);
}

//********************************************************
cursor_colser::cursor_colser(const odbc_raii_statement& h, bool b) : h_(h), close_(b)
{   }

cursor_colser::~cursor_colser()
{
    if (close_)
        h_.invoke(
            [](HSTMT x) { return ::SQLCloseCursor(x); }
        );
}

//********************************************************

odbc_set::odbc_set(const tstring& connectName, __int32& myNo) : pNo{ &myNo }
{
    env.AllocHandle();
    con.AllocHandle(env);
    st.AllocHandle(connectName, con);
}

odbc_set::~odbc_set()
{
    if (!too_late_to_destruct)  *pNo = -1;
}

odbc_raii_statement&  odbc_set::stmt()
{
    return st;
}

//********************************************************

RETCODE execDirect(const tstring& sql_expr, const odbc_raii_statement& stmt)
{
    SQLTCHAR* sql = const_cast<SQLTCHAR*>(static_cast<const SQLTCHAR*>(sql_expr.c_str()));
    return stmt.invoke(
        [=](HSTMT x) { return ::SQLExecDirect(x, sql, SQL_NTS); }
    );
}
//********************************************************

namespace 
{
    RETCODE
        do_SQLDescribeCol(  const odbc_raii_statement&  stmt            ,
                            UWORD                       colnumber       ,
                            TCHAR*                      pcolname        ,
                            SWORD                       sizeofcolname   ,
                            SQLSMALLINT*                pcolnamelen     ,
                            SQLSMALLINT*                pcoltype        ,
                            SQLULEN*                    pcollen         ,
                            SQLSMALLINT*                pscale          ,
                            SQLSMALLINT*                pnullable       )
    {
        auto expr = [=](HSTMT x){ return ::SQLDescribeCol( x,
                                                        colnumber,
                                                    pcolname,
                                                sizeofcolname,
                                            pcolnamelen,
                                        pcoltype,
                                    pcollen,
                                pscale,
                            pnullable);
                        };
        return stmt.invoke(expr);
    }

    RETCODE
        do_SQLBindCol(  const odbc_raii_statement&  stmt        ,
                        UWORD                       colnumber   ,
                        UCHAR*                      datai       ,
                        std::size_t                 bufsize     ,
                        SQLLEN*                     datastrleni )
    {
        return stmt.invoke(
            [=](HSTMT x) { return ::SQLBindCol(x, colnumber, SQL_C_CHAR, datai, bufsize, datastrleni); }
        );
    }
}

//*************************************************

tstring getTypeStr(SQLSMALLINT type)
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
