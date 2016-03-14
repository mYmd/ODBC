//vbODBC.cpp
//Copyright (c) 2015 mmYYmmdd

//******************************
//  terminateODBC
//  initODBC
//	selectODBC
//	ODBC_columnAttribute
//	execODBC
//******************************
#include "stdafx.h"
#include "odbcResource.hpp"
#include <memory>
#include <vector>
#include <OleAuto.h>

namespace   {
    using tstring = odbc_raii_statement::tstring;

    struct odbc_set  {  
        odbc_raii_env        env;
        odbc_raii_connect    con;
        odbc_raii_statement  st;
        odbc_set(const tstring& connectName) : env{}, con{}
        {
            env.AllocHandle();
            con.AllocHandle(env);
            st.AllocHandle(connectName, con);
        }
    };

    using pODBCStmt = std::unique_ptr<odbc_set>;
    std::vector<pODBCStmt>                  vODBCStmt;

    tstring getTypeStr(SQLSMALLINT);
    VARIANT makeVariantFromSQLType(SQLSMALLINT, LPCOLESTR);
    auto selectODBC_result(__int32, VARIANT*, std::vector<SQLSMALLINT>&, __int32)
        ->odbc_raii_select::result_type;
    BSTR getBSTR(VARIANT* expr);
}

__int32 __stdcall terminateODBC(__int32 myNo)
{
    if ( 0 <= myNo && myNo < vODBCStmt.size() )
        vODBCStmt[myNo].reset();
    return 0;
}

__int32 __stdcall initODBC(__int32 myNo, VARIANT* rawStr)
{
    BSTR bstr = getBSTR(rawStr);
    if (!bstr )                     return -1;
    tstring connectName{bstr};
    try
    {
        if ( 0 <= myNo && myNo < vODBCStmt.size() )
        {
            vODBCStmt[myNo] = std::make_unique<odbc_set>(connectName);
        }
        else
        {
            vODBCStmt.push_back(std::make_unique<odbc_set>(connectName));
            myNo = static_cast<int>(vODBCStmt.size() - 1);
        }
    }
    catch ( RETCODE )
    {
        terminateODBC(myNo);
        return -1;
    }
    return myNo;
}

VARIANT __stdcall selectODBC(__int32 myNo, VARIANT* SQL, __int32 timeOutSec)
{
    VARIANT ret;
    ::VariantInit(&ret);
    std::vector<SQLSMALLINT> coltype;
    auto result = selectODBC_result(myNo, SQL, coltype, timeOutSec);
    if ( result.empty() )   return ret;
    std::size_t const row = result.size();
    std::size_t const col = result[0].size();
    if ( 0==col )           return ret;
    SAFEARRAYBOUND rgb[2] = { { static_cast<ULONG>(result.size()), 0 },
                              { static_cast<ULONG>(result[0].size()), 0 } };
    SAFEARRAY* pArray = ::SafeArrayCreate(VT_VARIANT, 2, rgb);
    auto const elemsize = ::SafeArrayGetElemsize(pArray);
    char* it = nullptr;
    ::SafeArrayAccessData(pArray, reinterpret_cast<void**>(&it));
    if (!it)
    {
        ::SafeArrayUnaccessData(pArray);
        return ret;
    }
    for (std::size_t i = 0; i < row; ++i)
    {
        for (std::size_t j = 0; j < col; ++j)
        {
            tstring const& str = result[i][j];
            TCHAR const* p = str.empty() ? 0 : &str[0];
            VARIANT elem = makeVariantFromSQLType(coltype[j], p);
            auto const distance = result.size() * j + i;
            std::swap(*reinterpret_cast<VARIANT*>(it + distance * elemsize), elem);
            ::VariantClear(&elem);
        }
    }
    ::SafeArrayUnaccessData(pArray);
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = pArray;
    return ret;
}

VARIANT __stdcall selectODBC_zip(__int32 myNo, VARIANT* SQL, __int32 timeOutSec)
{
    VARIANT ret;
    ::VariantInit(&ret);
    std::vector<SQLSMALLINT> coltype;
    auto result = selectODBC_result(myNo, SQL, coltype, timeOutSec);
    if ( result.empty() )   return ret;
    std::size_t const row = result.size();
    std::size_t const col = result[0].size();
    if ( 0==col )           return ret;
    SAFEARRAYBOUND rgb[2] = { { static_cast<ULONG>(result.size()), 0 },
                              { static_cast<ULONG>(result[0].size()), 0 } };
    SAFEARRAY* pArray1 = ::SafeArrayCreate(VT_VARIANT, 1, rgb + 0);
    auto const elemsize1 = ::SafeArrayGetElemsize(pArray1);
    char* it1 = nullptr;
    ::SafeArrayAccessData(pArray1, reinterpret_cast<void**>(&it1));
    if (!it1)
    {
        ::SafeArrayUnaccessData(pArray1);
        return ret;
    }
    for (std::size_t i = 0; i < row; ++i)
    {
        SAFEARRAY* pArray2 = ::SafeArrayCreate(VT_VARIANT, 1, rgb + 1);
        auto const elemsize2 = ::SafeArrayGetElemsize(pArray2);
        char* it2 = nullptr;
        ::SafeArrayAccessData(pArray2, reinterpret_cast<void**>(&it2));
        for (std::size_t j = 0; j < col; ++j)
        {
            tstring const& str = result[i][j];
            TCHAR const* p = str.empty() ? 0 : &str[0];
            VARIANT elem = makeVariantFromSQLType(coltype[j], p);
            std::swap(*reinterpret_cast<VARIANT*>(it2 + j * elemsize2), elem);
            ::VariantClear(&elem);
        }
        ::SafeArrayUnaccessData(pArray2);
        reinterpret_cast<VARIANT*>(it1 + i * elemsize1)->vt = VT_ARRAY | VT_VARIANT;
        reinterpret_cast<VARIANT*>(it1 + i * elemsize1)->parray = pArray2;
    }
    ::SafeArrayUnaccessData(pArray1);
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = pArray1;
    return ret;
}

VARIANT __stdcall columnAttributes(__int32 myNo, VARIANT* SQL)
{
    VARIANT ret;
    ::VariantInit(&ret);
    BSTR bstr = getBSTR(SQL);
    if ( !bstr || myNo < 0 || vODBCStmt.size() <= myNo )
        return ret;
    using column_name_type = odbc_raii_select::column_name_type;
    std::vector<column_name_type>   colname;
    std::vector<SQLSMALLINT>        colnamelen;
    std::vector<SQLULEN>            collen;
    std::vector<SQLSMALLINT>        nullable;
    std::vector<SQLSMALLINT>        coltype;
    std::vector<SQLSMALLINT>        scale;
    std::vector<SQLLEN>             datastrlen;
    SQLSMALLINT     nresultcols = 0;
    try {
        odbc_raii_select    odbcSelect;
        cursor_colser       c_closer(vODBCStmt[myNo]->st);
        nresultcols = odbcSelect.columnAttribute(tstring(bstr),
                                                vODBCStmt[myNo]->st,
                                                colname,
                                                colnamelen,
                                                collen,
                                                nullable,
                                                coltype,
                                                scale,
                                                datastrlen,
                                                0);
    }
    catch (RETCODE )
    {
        return ret;
    }
    if ( nresultcols == 0 )         return ret;
    SAFEARRAYBOUND rgb[2] = { { nresultcols, 0 },  { 2, 0 } };
    SAFEARRAY* pArray = ::SafeArrayCreate(VT_VARIANT, 2, rgb);
    auto const elemsize = ::SafeArrayGetElemsize(pArray);
    char* it = nullptr;
    ::SafeArrayAccessData(pArray, reinterpret_cast<void**>(&it));
    if (!it)
    {
        ::SafeArrayUnaccessData(pArray);
        return ret;
    }
    for ( SQLSMALLINT i = 0; i < nresultcols; ++i )
    {
        {
            VARIANT elem;
            ::VariantInit(&elem);
            elem.vt = VT_BSTR;
            TCHAR const* p = colname[i].data();
            elem.bstrVal = SysAllocString(p);
            ::VariantCopy(reinterpret_cast<VARIANT*>(it + i * elemsize), &elem);
            ::VariantClear(&elem);
        }
        {
            VARIANT elem;
            ::VariantInit(&elem);
            elem.vt = VT_BSTR;
            tstring const str = getTypeStr(coltype[i]);
            TCHAR const* p = str.empty() ? 0 : &str[0];
            elem.bstrVal = SysAllocString(p);
            ::VariantCopy(reinterpret_cast<VARIANT*>(it + (nresultcols + i) * elemsize), &elem);
            ::VariantClear(&elem);
        }
    }
    ::SafeArrayUnaccessData(pArray);
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = pArray;
    return ret;
}

__int32 __stdcall execODBC(__int32 myNo, VARIANT* SQLs)
{
    if ( myNo < 0 || vODBCStmt.size() <= myNo )         return 0;
    if ( !SQLs ||  0 == (VT_ARRAY & SQLs->vt ) )        return 0;
    SAFEARRAY* pArray = ( 0 == (VT_BYREF & SQLs->vt) )?  (SQLs->parray): (*SQLs->pparray);
    if ( !pArray || 1 != ::SafeArrayGetDim(pArray) )    return 0;
    SAFEARRAYBOUND bounds = {1,0};   //要素数、LBound
    {
        ::SafeArrayGetLBound(pArray, 1, &bounds.lLbound);
        LONG ub = 0;
        ::SafeArrayGetUBound(pArray, 1, &ub);
        bounds.cElements = 1 + ub - bounds.lLbound;
    }
    __int32 count = 0;
    odbc_raii_select    odbcSelect;
    cursor_colser       c_close(vODBCStmt[myNo]->st);
    VARIANT elem;
    ::VariantInit(&elem);
    for ( ULONG i = 0; i < bounds.cElements; ++i )
    {
        LONG index = static_cast<LONG>(i) + bounds.lLbound;
        ::SafeArrayGetElement(pArray, &index, &elem);
        if (elem.vt == VT_BSTR && elem.bstrVal )
        {
            auto const rc = odbcSelect.execDirect(tstring(elem.bstrVal), vODBCStmt[myNo]->st);
            if ( rc == SQL_SUCCESS || rc == SQL_SUCCESS_WITH_INFO )     ++count;
        }
        ::VariantClear(&elem);
    }
    return count;
}

namespace   {
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
        case SQL_TYPE_TIMESTAMP:    ret = tstring(_T("TYPE_TIMESTAMP"));break;
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

    VARIANT makeVariantFromSQLType(SQLSMALLINT type, LPCOLESTR strln)
    {
        VARIANT ret;
        ::VariantInit(&ret);
        if ( !strln )
        {
            ret.vt = VT_NULL;
            return ret;
        }
        switch (type)
        {
        case SQL_CHAR:
        case SQL_VARCHAR:
        case SQL_LONGVARCHAR:
        case SQL_WCHAR:
        case SQL_WVARCHAR:
        case SQL_WLONGVARCHAR:
        case SQL_BINARY:
        case SQL_VARBINARY:
        case SQL_LONGVARBINARY:
        {
            ret.vt = VT_BSTR;
            ret.bstrVal = SysAllocString(strln);
            return ret;
        }
        case SQL_SMALLINT:
        case SQL_INTEGER:
        case SQL_BIT:
        case SQL_TINYINT:
        {
            long lOut;
            auto const vdr = VarI4FromStr(strln, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &lOut);
            ret.vt = VT_I4;
            ret.lVal = lOut;
            return ret;
        }
        case SQL_BIGINT:
        {
            LONG64  i64Out;
            auto const vdr = VarI8FromStr(strln, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &i64Out);
            ret.vt = VT_I8;
            ret.llVal = i64Out;
            return ret;
        }
        case SQL_NUMERIC:
        case SQL_DECIMAL:
        case SQL_FLOAT:
        case SQL_REAL:
        case SQL_DOUBLE:
        {
            double dOut;
            auto const vdr = VarR8FromStr(strln, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &dOut);
            ret.vt = VT_R8;
            ret.dblVal = dOut;
            return ret;
        }
        case SQL_TYPE_DATE:
        case SQL_TYPE_TIME:
        case SQL_TYPE_TIMESTAMP:
        {
            OLECHAR strln2[] = _T("2001-01-01 00:00:00");
            auto p = strln;
            auto q = strln2;
            while ( *p != _T('\0') && *p != _T('.') && *q != _T('\0') )
                *q++ = *p++;
            *q = _T('\0');
            DATE dOut;
            auto const vdr = VarDateFromStr(strln2, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &dOut);
            ret.vt = VT_DATE;
            ret.date = dOut;
            return ret;
        }
        default:
            ret.vt = VT_NULL;
            return ret;
        }
    }

    auto selectODBC_result(__int32 myNo, VARIANT* SQL, std::vector<SQLSMALLINT>& coltype, __int32 timeOutSec)
        ->odbc_raii_select::result_type
    {
        odbc_raii_select::result_type result;
        BSTR bstr = getBSTR(SQL);
        if ( !bstr || myNo < 0 || vODBCStmt.size() <= myNo )
            return result;
        odbc_raii_select    odbcSelect;
        cursor_colser       c_close(vODBCStmt[myNo]->st);
        try {
            result = odbcSelect.select(  timeOutSec,
                                         tstring(bstr),
                                         vODBCStmt[myNo]->st,
                                         nullptr, nullptr, nullptr, nullptr,
                                         &coltype,
                                         nullptr, nullptr);
        }
        catch (RETCODE )
        {
            result.erase(result.begin(), result.end());
        }
        return result;
    }

    BSTR getBSTR(VARIANT* expr)
    {
        if ( !expr )
            return nullptr;
        else if ( expr->vt & VT_BYREF )
            return ( (expr->vt & VT_BSTR) && expr->pbstrVal )? *expr->pbstrVal : nullptr;
        else
            return ( (expr->vt & VT_BSTR) && expr->bstrVal )?   expr->bstrVal  : nullptr;
    }

}
