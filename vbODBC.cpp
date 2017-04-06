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

    void selectODBC_rcWise_imple(VARIANT&, odbc_raii_select::result_type const&, std::vector<SQLSMALLINT> const&, bool);

    template <typename F>
    VARIANT catalogValue(F&&, __int32, VARIANT*, VARIANT*, SQLUSMALLINT);

    class safearrayRAII    {
        SAFEARRAY* pArray;
    public:
        safearrayRAII(SAFEARRAY* p) : pArray(p) {}
        ~safearrayRAII()            {   ::SafeArrayUnaccessData(pArray);    }
        SAFEARRAY* get() const      {   return pArray;  }
    };

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
    safearrayRAII pArray(::SafeArrayCreate(VT_VARIANT, 2, rgb));
    auto const elemsize = ::SafeArrayGetElemsize(pArray.get());
    char* it = nullptr;
    ::SafeArrayAccessData(pArray.get(), reinterpret_cast<void**>(&it));
    if (!it)
        return ret;
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
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = pArray.get();
    return ret;
}

VARIANT __stdcall selectODBC_rowWise(__int32 myNo, VARIANT* SQL, __int32 timeOutSec)
{
    VARIANT ret;
    ::VariantInit(&ret);
    std::vector<SQLSMALLINT> coltype;
    auto result = selectODBC_result(myNo, SQL, coltype, timeOutSec);
    selectODBC_rcWise_imple(ret, result, coltype, true);
    return ret;
}

VARIANT __stdcall selectODBC_columnWise(__int32 myNo, VARIANT* SQL, __int32 timeOutSec)
{
    VARIANT ret;
    ::VariantInit(&ret);
    std::vector<SQLSMALLINT> coltype;
    auto result = selectODBC_result(myNo, SQL, coltype, timeOutSec);
    selectODBC_rcWise_imple(ret, result, coltype, false);
    return ret;
}

VARIANT __stdcall columnAttributes(__int32 myNo, VARIANT* SQL, __int32 getNullable)
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
    SAFEARRAYBOUND rgb[2] = { {static_cast<ULONG>(nresultcols), 0}, {(getNullable? 5U: 4U), 0} };
    safearrayRAII pArray(::SafeArrayCreate(VT_VARIANT, 2, rgb));
    auto const elemsize = ::SafeArrayGetElemsize(pArray.get());
    char* it = nullptr;
    ::SafeArrayAccessData(pArray.get(), reinterpret_cast<void**>(&it));
    if (!it)
        return ret;
    for ( SQLSMALLINT i = 0; i < nresultcols; ++i )
    {
        {
            VARIANT& elem = *reinterpret_cast<VARIANT*>(it + i * elemsize);
            elem.vt = VT_BSTR;
            elem.bstrVal = SysAllocString(colname[i].data());
        }
        {
            VARIANT& elem = *reinterpret_cast<VARIANT*>(it + (nresultcols + i) * elemsize);
            elem.vt = VT_BSTR;
            tstring const str = getTypeStr(coltype[i]);
            TCHAR const* p = str.empty() ? 0 : &str[0];
            elem.bstrVal = SysAllocString(p);
        }
        reinterpret_cast<VARIANT*>(it + (2*nresultcols + i) * elemsize)->vt = VT_I4;
        reinterpret_cast<VARIANT*>(it + (2*nresultcols + i) * elemsize)->lVal = static_cast<LONG>(collen[i]);
        reinterpret_cast<VARIANT*>(it + (3*nresultcols + i) * elemsize)->vt = VT_I4;
        reinterpret_cast<VARIANT*>(it + (3*nresultcols + i) * elemsize)->lVal = scale[i];
        if ( getNullable )
        {
            reinterpret_cast<VARIANT*>(it + (4*nresultcols + i) * elemsize)->vt = VT_I4;
            reinterpret_cast<VARIANT*>(it + (4*nresultcols + i) * elemsize)->lVal = (nullable[i]? -1: 0);
        }
    }
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = pArray.get();
    return ret;
}

VARIANT __stdcall execODBC(__int32 myNo, VARIANT* SQLs)
{
    VARIANT ret;
    ::VariantInit(&ret);
    if ( myNo < 0 || vODBCStmt.size() <= myNo )         return ret;
    if ( !SQLs ||  0 == (VT_ARRAY & SQLs->vt ) )        return ret;
    SAFEARRAY* pArray = ( 0 == (VT_BYREF & SQLs->vt) )?  (SQLs->parray): (*SQLs->pparray);
    if ( !pArray || 1 != ::SafeArrayGetDim(pArray) )    return ret;
    SAFEARRAYBOUND bounds = {1,0};   //要素数、LBound
    {
        ::SafeArrayGetLBound(pArray, 1, &bounds.lLbound);
        LONG ub = 0;
        ::SafeArrayGetUBound(pArray, 1, &ub);
        bounds.cElements = 1 + ub - bounds.lLbound;
    }
    odbc_raii_select    odbcSelect;
    cursor_colser       c_close(vODBCStmt[myNo]->st);
    VARIANT elem;
    ::VariantInit(&elem);
    std::vector<LONG> errorNo;
    for ( ULONG i = 0; i < bounds.cElements; ++i )
    {
        LONG index = static_cast<LONG>(i) + bounds.lLbound;
        ::SafeArrayGetElement(pArray, &index, &elem);
        if (elem.vt == VT_BSTR && elem.bstrVal )
        {
            auto const rc = odbcSelect.execDirect(tstring(elem.bstrVal), vODBCStmt[myNo]->st);
            if ( rc != SQL_SUCCESS && rc != SQL_SUCCESS_WITH_INFO )
                errorNo.push_back(index);
        }
        ::VariantClear(&elem);
    }
    if ( errorNo.size() )
    {
        SAFEARRAYBOUND rgb = { static_cast<ULONG>(errorNo.size()), 0 };
        safearrayRAII pNo(::SafeArrayCreate(VT_VARIANT, 1, &rgb));
        auto const elemsize = ::SafeArrayGetElemsize(pNo.get());
        char* it = nullptr;
        ::SafeArrayAccessData(pNo.get(), reinterpret_cast<void**>(&it));
        for ( auto i = 0; i < errorNo.size(); ++i )
        {
            reinterpret_cast<VARIANT*>(it + i*elemsize)->vt = VT_I4;
            reinterpret_cast<VARIANT*>(it + i*elemsize)->lVal = errorNo[i];
        }
        ret.vt = VT_ARRAY | VT_VARIANT;
        ret.parray = pNo.get();
    }
    return ret;
}

VARIANT __stdcall table_list_all(__int32 myNo, VARIANT* schemaName)
{
    struct table_func_t {       //VC2013対策
        SQLTCHAR* scName;   SQLSMALLINT scLen;
        SQLRETURN operator()(HSTMT x) const
        {
            return ::SQLTables(x, NULL, SQL_NTS, scName, scLen, NULL, SQL_NTS, NULL, SQL_NTS);
        }
    };
    auto table_func = [](SQLTCHAR* scName, SQLSMALLINT scLen, SQLTCHAR* Dummy, SQLSMALLINT dummy)
    {
        //return [=](HSTMT x) {     //VC2013ではNG
        //    return ::SQLTables(x, NULL, SQL_NTS, scName, scLen, NULL, SQL_NTS, NULL, SQL_NTS);
        //};
        return table_func_t{ scName, scLen };
    };
    VARIANT table_name = catalogValue(table_func, myNo, schemaName, schemaName, 3);     //TABLE_NAME
    VARIANT type_name  = catalogValue(table_func, myNo, schemaName, schemaName, 4);     //TABLE_TYPE
    SAFEARRAYBOUND rgb = { ULONG{2}, 0 };
    safearrayRAII pArray(::SafeArrayCreate(VT_VARIANT, 1, &rgb));
    char* it = nullptr;
    ::SafeArrayAccessData(pArray.get(), reinterpret_cast<void**>(&it));
    VARIANT ret;
    ::VariantInit(&ret);
    if ( !it )  return ret;
    auto const elemsize = ::SafeArrayGetElemsize(pArray.get());
    std::swap(*reinterpret_cast<VARIANT*>(it), table_name);
    std::swap(*reinterpret_cast<VARIANT*>(it + elemsize), type_name);
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = pArray.get();
    return ret;
}

// https://www.ibm.com/support/knowledgecenter/ja/SSEPEK_11.0.0/odbc/src/tpc/db2z_fnprimarykeys.html#db2z_fnpkey__bknetbprkey
// https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlprimarykeys-function
VARIANT __stdcall columnAttributes_all(__int32 myNo, VARIANT* schemaName, VARIANT* tableName)
{
    struct column_func_t   {       //VC2013対策
        SQLTCHAR* scName, *tbName;  SQLSMALLINT scLen, tbLen;
        SQLRETURN operator()(HSTMT x) const
        {   return ::SQLColumns(x, NULL, SQL_NTS, scName, scLen, tbName, tbLen, NULL, SQL_NTS); }
    };
    struct primarykeys_func_t {     //VC2013対策
        SQLTCHAR* scName, *tbName;  SQLSMALLINT scLen, tbLen;
        SQLRETURN operator()(HSTMT x) const
        {   return ::SQLPrimaryKeys(x, NULL, SQL_NTS, scName, scLen, tbName, tbLen);    }
    };
    auto column_func = [](SQLTCHAR* scName, SQLSMALLINT scLen, SQLTCHAR* tbName, SQLSMALLINT tbLen)
    {
        //return [=](HSTMT x) {     //VC2013ではNG
        //    return ::SQLColumns(x, NULL, SQL_NTS, scName, scLen, tbName, tbLen, NULL, SQL_NTS);
        //};
        return column_func_t{ scName, tbName, scLen, tbLen };
    };
    auto primarykeys_func = [](SQLTCHAR* scName, SQLSMALLINT scLen, SQLTCHAR* tbName, SQLSMALLINT tbLen)
    {
        //return [=](HSTMT x) {     //VC2013ではNG
        //    return ::SQLColumns(x, NULL, SQL_NTS, scName, scLen, tbName, tbLen, NULL, SQL_NTS);
        //};
        return primarykeys_func_t{ scName, tbName, scLen, tbLen };
    };
    VARIANT column_name         = catalogValue(column_func, myNo, schemaName, tableName, 4);        // COLUMN_NAME
    VARIANT type_name           = catalogValue(column_func, myNo, schemaName, tableName, 6);        // TYPE_NAME
    VARIANT column_size         = catalogValue(column_func, myNo, schemaName, tableName, 7);        // COLUMN_SIZE
    VARIANT is_nullable         = catalogValue(column_func, myNo, schemaName, tableName, 11);       // IS_NULLABLE
    VARIANT ordinal_position    = catalogValue(column_func, myNo, schemaName, tableName, 17);       // ORDINAL_POSITION
    VARIANT primarykeys         = catalogValue(primarykeys_func, myNo, schemaName, tableName, 4);   // COLUMN_NAME
    SAFEARRAYBOUND rgb = { ULONG{6}, 0 };
    safearrayRAII pArray(::SafeArrayCreate(VT_VARIANT, 1, &rgb));
    char* it = nullptr;
    ::SafeArrayAccessData(pArray.get(), reinterpret_cast<void**>(&it));
    VARIANT ret;
    ::VariantInit(&ret);
    if (!it)
        return ret;
    auto const elemsize = ::SafeArrayGetElemsize(pArray.get());
    VARIANT* pvec[6] = { &column_name, &type_name, &column_size, &is_nullable, &ordinal_position, &primarykeys };
    for (std::size_t i = 0; i < 6; ++i)
        std::swap(*reinterpret_cast<VARIANT*>(it + i * elemsize), *pvec[i]);
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = pArray.get();
    return ret;
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
        if (!expr)
            return nullptr;
        else if (expr->vt & VT_BYREF)
            return ((expr->vt & VT_BSTR) && expr->pbstrVal) ? *expr->pbstrVal : nullptr;
        else
            return ((expr->vt & VT_BSTR) && expr->bstrVal) ? expr->bstrVal : nullptr;
    }

    void selectODBC_rcWise_imple(VARIANT&                               ret,
                                odbc_raii_select::result_type const&    result,
                                std::vector<SQLSMALLINT> const&         coltype,
                                bool                                    rc = true)
    {
        if (0 == result.size() || 0 == result[0].size())      return;
        std::size_t const outerSize = rc ? result.size() : result[0].size();
        std::size_t const innerSize = rc ? result[0].size() : result.size();
        SAFEARRAYBOUND outerBound = { static_cast<ULONG>(outerSize), 0 };
        SAFEARRAYBOUND innerBound = { static_cast<ULONG>(innerSize), 0 };
        safearrayRAII outerArray(::SafeArrayCreate(VT_VARIANT, 1, &outerBound));  //
        auto const outerElemSize = ::SafeArrayGetElemsize(outerArray.get());
        char* outerBegin = nullptr;
        ::SafeArrayAccessData(outerArray.get(), reinterpret_cast<void**>(&outerBegin));
        if (!outerBegin)
            return;
        for (std::size_t outerIter = 0; outerIter < outerSize; ++outerIter)
        {
            safearrayRAII innerArray(::SafeArrayCreate(VT_VARIANT, 1, &innerBound));
            auto const innerElemSize = ::SafeArrayGetElemsize(innerArray.get());
            char* innerBegin = nullptr;
            ::SafeArrayAccessData(innerArray.get(), reinterpret_cast<void**>(&innerBegin));
            for (std::size_t innerIter = 0; innerIter < innerSize; ++innerIter)
            {
                tstring const& str = rc ?
                    result[outerIter][innerIter] :
                    result[innerIter][outerIter];
                TCHAR const* p = str.empty() ? 0 : &str[0];
                VARIANT elem = makeVariantFromSQLType(coltype[rc? innerIter: outerIter], p);
                std::swap(*reinterpret_cast<VARIANT*>(innerBegin + innerIter * innerElemSize), elem);
                ::VariantClear(&elem);
            }
            reinterpret_cast<VARIANT*>(outerBegin + outerIter * outerElemSize)->vt = VT_ARRAY | VT_VARIANT;
            reinterpret_cast<VARIANT*>(outerBegin + outerIter * outerElemSize)->parray = innerArray.get();
        }
        ret.vt = VT_ARRAY | VT_VARIANT;
        ret.parray = outerArray.get();
    }

    template <typename F>
    VARIANT catalogValue(F&& func, __int32 myNo, VARIANT* schemaName, VARIANT* tableName, SQLUSMALLINT ColumnNumber)
    {
        VARIANT ret;
        ::VariantInit(&ret);
        BSTR schema_name_b = getBSTR(schemaName);
        BSTR table_name_b  = getBSTR(tableName);
        if (!schema_name_b || !table_name_b || myNo < 0 || vODBCStmt.size() <= myNo)
            return ret;
        auto& st = vODBCStmt[myNo]->st;
        cursor_colser   c_closer(st);
        {
            tstring schema_name_t(schema_name_b);
            tstring table_name_t(table_name_b);
            SQLTCHAR* schema_name = const_cast<SQLTCHAR*>(static_cast<const SQLTCHAR*>(schema_name_t.c_str()));
            auto schema_len = static_cast<SQLSMALLINT>(schema_name_t.length());
            if ( schema_len == 0 )      schema_name = NULL;
            SQLTCHAR* table_name = const_cast<SQLTCHAR*>(static_cast<const SQLTCHAR*>(table_name_t.c_str()));
            auto func_expr = std::forward<F>(func)( schema_name,
                                                    schema_len,
                                                    table_name,
                                                    static_cast<SQLSMALLINT>(table_name_t.length()) );
            auto result = st.invoke(func_expr);
            if (SQL_SUCCESS != result)
                return ret;
        }
        SQLSMALLINT nresultcols = 0;
        {
            SQLSMALLINT* pl = &nresultcols;
            auto const expr = [=](HSTMT x){return ::SQLNumResultCols(x, pl); };
            auto result = st.invoke(expr);
        }
        SQLCHAR  rgbValue[odbc_raii_select::ColumnNameLen];
        SQLLEN   pcbValue;
        {
            auto p_rgbValue = static_cast<SQLPOINTER>(rgbValue);
            auto p_pcbValue = &pcbValue;
            for ( auto j = 0; j < nresultcols; ++j )
            {
                if ( j == ColumnNumber )    continue;
                auto SQLBindCol_expr_ = [=](HSTMT x) {
                    return ::SQLBindCol(x, j, SQL_C_CHAR, NULL, 0, NULL);
                };
                auto result = st.invoke(SQLBindCol_expr_);
            }
            auto SQLBindCol_expr = [=](HSTMT x) {
                return ::SQLBindCol(x,
                            ColumnNumber,      //COLUMN_NAME
                            SQL_C_CHAR,
                            p_rgbValue,
                            odbc_raii_select::ColumnNameLen,
                            p_pcbValue);
            };
            auto result = st.invoke(SQLBindCol_expr);
            if (SQL_SUCCESS != result)
                return ret;
        }
        auto SQLFetch_expr = [=](HSTMT x) { return ::SQLFetch(x); };
        std::vector<VARIANT> vec;
        while (true)
        {
            auto fetch_result = st.invoke(SQLFetch_expr);
            if ((SQL_SUCCESS != fetch_result) && (SQL_SUCCESS_WITH_INFO != fetch_result))
                break;
            TCHAR tcharBuffer[odbc_raii_select::ColumnNameLen];
            int mb = ::MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, (LPCSTR)rgbValue, -1, tcharBuffer, odbc_raii_select::ColumnNameLen);
            tstring const str(tcharBuffer);
            TCHAR const* p = str.empty() ? 0 : &str[0];
            vec.push_back(makeVariantFromSQLType(SQL_CHAR, p));
        }
        SAFEARRAYBOUND rgb = { static_cast<ULONG>(vec.size()), 0 };
        safearrayRAII pArray(::SafeArrayCreate(VT_VARIANT, 1, &rgb));
        char* it = nullptr;
        ::SafeArrayAccessData(pArray.get(), reinterpret_cast<void**>(&it));
        if (!it)
            return ret;
        auto const elemsize = ::SafeArrayGetElemsize(pArray.get());
        for (std::size_t i = 0; i < vec.size(); ++i)
            std::swap(*reinterpret_cast<VARIANT*>(it + i * elemsize), vec[i]);
        ret.vt = VT_ARRAY | VT_VARIANT;
        ret.parray = pArray.get();
        return ret;
    }

}
