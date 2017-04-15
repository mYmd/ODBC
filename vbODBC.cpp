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


namespace {

    std::vector<std::unique_ptr<odbc_set>>  vODBCStmt;

    VARIANT makeVariantFromSQLType(SQLSMALLINT, LPCOLESTR);

    auto selectODBC_result(__int32, VARIANT*, std::vector<SQLSMALLINT>&, __int32)
        ->odbc_raii_select::result_type;

    BSTR getBSTR(VARIANT* expr);

    void selectODBC_rcWise_imple(VARIANT&, odbc_raii_select::result_type const&, std::vector<SQLSMALLINT> const&, bool);

    class safearrayRAII {
        SAFEARRAY* pArray;
    public:
        safearrayRAII(SAFEARRAY* p) : pArray(p) {}
        ~safearrayRAII() { ::SafeArrayUnaccessData(pArray); }
        SAFEARRAY* get() const { return pArray; }
    };

    VARIANT vec2VArray(std::vector<VARIANT>& vec);
}

//----------------------------------------------------------------------

void __stdcall terminateODBC(__int32 myNo)
{
    if ( 0 <= myNo && myNo < vODBCStmt.size() )
        vODBCStmt[myNo].reset();
}

void __stdcall terminateODBC_all()
{
    vODBCStmt.clear();
}

__int32 __stdcall initODBC(__int32& myNo, VARIANT* rawStr)
{
    BSTR bstr = getBSTR(rawStr);
    if (!bstr)                     return -1;
    tstring connectName{ bstr };
    try
    {
        if ( 0 <= myNo && myNo < vODBCStmt.size() )
        {
            auto tmp = myNo;
            vODBCStmt[myNo] = std::make_unique<odbc_set>(connectName, myNo);
            myNo = tmp;
        }
        else
        {
            vODBCStmt.push_back(std::make_unique<odbc_set>(connectName, myNo));
            myNo = static_cast<int>(vODBCStmt.size() - 1);
        }
    }
    catch (RETCODE)
    {
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
    if (result.empty())   return ret;
    std::size_t const row = result.size();
    std::size_t const col = result[0].size();
    if (0==col)           return ret;
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
        cursor_colser       c_closer(vODBCStmt[myNo]->stmt());
        nresultcols = odbcSelect.columnAttribute(tstring(bstr),
            vODBCStmt[myNo]->stmt(),
            colname,
            colnamelen,
            collen,
            nullable,
            coltype,
            scale,
            datastrlen,
            0);
    }
    catch (RETCODE)
    {
        return ret;
    }
    if (nresultcols == 0)         return ret;
    SAFEARRAYBOUND rgb[2] = { { static_cast<ULONG>(nresultcols), 0 },{ (getNullable ? 5U : 4U), 0 } };
    safearrayRAII pArray(::SafeArrayCreate(VT_VARIANT, 2, rgb));
    auto const elemsize = ::SafeArrayGetElemsize(pArray.get());
    char* it = nullptr;
    ::SafeArrayAccessData(pArray.get(), reinterpret_cast<void**>(&it));
    if (!it)
        return ret;
    for (SQLSMALLINT i = 0; i < nresultcols; ++i)
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
        if (getNullable)
        {
            reinterpret_cast<VARIANT*>(it + (4*nresultcols + i) * elemsize)->vt = VT_I4;
            reinterpret_cast<VARIANT*>(it + (4*nresultcols + i) * elemsize)->lVal = (nullable[i] ? -1 : 0);
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
    if ( myNo < 0 || vODBCStmt.size() <= myNo )    return ret;
    if (!SQLs ||  0 == (VT_ARRAY & SQLs->vt))        return ret;
    SAFEARRAY* pArray = (0 == (VT_BYREF & SQLs->vt)) ? (SQLs->parray) : (*SQLs->pparray);
    if (!pArray || 1 != ::SafeArrayGetDim(pArray))    return ret;
    SAFEARRAYBOUND bounds = { 1,0 };   //要素数、LBound
    {
        ::SafeArrayGetLBound(pArray, 1, &bounds.lLbound);
        LONG ub = 0;
        ::SafeArrayGetUBound(pArray, 1, &ub);
        bounds.cElements = 1 + ub - bounds.lLbound;
    }
    odbc_raii_select    odbcSelect;
    cursor_colser       c_close(vODBCStmt[myNo]->stmt());
    VARIANT elem;
    ::VariantInit(&elem);
    std::vector<LONG> errorNo;
    for (ULONG i = 0; i < bounds.cElements; ++i)
    {
        LONG index = static_cast<LONG>(i) + bounds.lLbound;
        ::SafeArrayGetElement(pArray, &index, &elem);
        if (elem.vt == VT_BSTR && elem.bstrVal)
        {
            auto const rc = odbcSelect.execDirect(tstring(elem.bstrVal), vODBCStmt[myNo]->stmt());
            if (rc != SQL_SUCCESS && rc != SQL_SUCCESS_WITH_INFO)
                errorNo.push_back(index);
        }
        ::VariantClear(&elem);
    }
    if (errorNo.size())
    {
        SAFEARRAYBOUND rgb = { static_cast<ULONG>(errorNo.size()), 0 };
        safearrayRAII pNo(::SafeArrayCreate(VT_VARIANT, 1, &rgb));
        auto const elemsize = ::SafeArrayGetElemsize(pNo.get());
        char* it = nullptr;
        ::SafeArrayAccessData(pNo.get(), reinterpret_cast<void**>(&it));
        for (auto i = 0; i < errorNo.size(); ++i)
        {
            reinterpret_cast<VARIANT*>(it + i*elemsize)->vt = VT_I4;
            reinterpret_cast<VARIANT*>(it + i*elemsize)->lVal = errorNo[i];
        }
        ret.vt = VT_ARRAY | VT_VARIANT;
        ret.parray = pNo.get();
    }
    return ret;
}

// テーブル一覧
VARIANT __stdcall table_list_all(__int32 myNo, VARIANT* schemaName)
{
    VARIANT ret;
    ::VariantInit(&ret);
    BSTR schema_name_b = getBSTR(schemaName);
    if (!schema_name_b || myNo < 0 || vODBCStmt.size() <= myNo)
        return ret;
    tstring schema_name_t(schema_name_b);
    SQLTCHAR* schema_name = const_cast<SQLTCHAR*>(static_cast<const SQLTCHAR*>(schema_name_t.c_str()));
    auto schema_len = static_cast<SQLSMALLINT>(schema_name_t.length());
    if (schema_len == 0)      schema_name = NULL;
    auto table_func = [=](HSTMT x)  {
        return ::SQLTables(x, NULL, SQL_NTS, schema_name, schema_len, NULL, SQL_NTS, NULL, SQL_NTS);
    };
    auto const& st = vODBCStmt[myNo]->stmt();
    std::vector<VARIANT> vec;
    auto push_back_func = [&](TCHAR const* p) {
        vec.push_back(makeVariantFromSQLType(SQL_CHAR, p));
    };
    catalogValue(table_func, st, 2, push_back_func);   //TABLE_SCHEM
    VARIANT schem_name = vec2VArray(vec);
    catalogValue(table_func, st, 3, push_back_func);   //TABLE_NAME
    VARIANT table_name = vec2VArray(vec);
    catalogValue(table_func, st, 4, push_back_func);    //TABLE_TYPE
    VARIANT type_name = vec2VArray(vec);
    vec.push_back(schem_name);
    vec.push_back(table_name);
    vec.push_back(type_name);
    return vec2VArray(vec);
}

// https://www.ibm.com/support/knowledgecenter/ja/SSEPEK_11.0.0/odbc/src/tpc/db2z_fnprimarykeys.html#db2z_fnpkey__bknetbprkey
// https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlprimarykeys-function
// テーブルにある全カラムの属性
VARIANT __stdcall columnAttributes_all(__int32 myNo, VARIANT* schemaName, VARIANT* tableName)
{
    VARIANT ret;
    ::VariantInit(&ret);
    BSTR schema_name_b{getBSTR(schemaName)}, table_Name_b{getBSTR(tableName)};
    if (!schema_name_b || !table_Name_b || myNo < 0 || vODBCStmt.size() <= myNo)
        return ret;
    tstring schema_name_t(schema_name_b), table_name_t(table_Name_b);
    SQLTCHAR* schema_name = const_cast<SQLTCHAR*>(static_cast<const SQLTCHAR*>(schema_name_t.c_str()));
    SQLTCHAR* table_Name  = const_cast<SQLTCHAR*>(static_cast<const SQLTCHAR*>(table_name_t.c_str()));
    auto schema_len = static_cast<SQLSMALLINT>(schema_name_t.length());
    auto table_len = static_cast<SQLSMALLINT>(table_name_t.length());
    if (schema_len == 0)      schema_name = NULL;
    auto column_func = [=](HSTMT x) {
        return ::SQLColumns(x, NULL, SQL_NTS, schema_name, schema_len, table_Name, table_len, NULL, SQL_NTS);
    };
    auto primarykeys_func = [=](HSTMT x) {
        return ::SQLPrimaryKeys(x, NULL, SQL_NTS, schema_name, schema_len, table_Name, table_len);
    };
    auto const& st = vODBCStmt[myNo]->stmt();
    std::vector<VARIANT> vec;
    auto push_back_func = [&](TCHAR const* p) {
        vec.push_back(makeVariantFromSQLType(SQL_CHAR, p));
    };
    catalogValue(column_func, st, 4, push_back_func);           // COLUMN_NAME
    VARIANT column_name = vec2VArray(vec);
    catalogValue(column_func, st, 6, push_back_func);           // TYPE_NAME
    VARIANT type_name = vec2VArray(vec);
    catalogValue(column_func, st, 7, push_back_func);           // COLUMN_SIZE
    VARIANT column_size = vec2VArray(vec);
    catalogValue(column_func, st, 11, push_back_func);          // IS_NULLABLE
    VARIANT is_nullable = vec2VArray(vec);
    catalogValue(column_func, st, 17, push_back_func);          // ORDINAL_POSITION
    VARIANT ordinal_position = vec2VArray(vec);
    catalogValue(primarykeys_func, st, 4, push_back_func);      // COLUMN_NAME
    VARIANT primarykeys = vec2VArray(vec);
    vec.push_back(column_name);
    vec.push_back(type_name);
    vec.push_back(column_size);
    vec.push_back(is_nullable);
    vec.push_back(ordinal_position);
    vec.push_back(primarykeys);
    return vec2VArray(vec);
}

namespace {

    VARIANT makeVariantFromSQLType(SQLSMALLINT type, LPCOLESTR strln)
    {
        VARIANT ret;
        ::VariantInit(&ret);
        if (!strln)
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
            while (*p != _T('\0') && *p != _T('.') && *q != _T('\0'))
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
        if (!bstr || myNo < 0 || vODBCStmt.size() <= myNo)
            return result;
        odbc_raii_select    odbcSelect;
        cursor_colser       c_close(vODBCStmt[myNo]->stmt());
        try {
            result = odbcSelect.select(timeOutSec,
                tstring(bstr),
                vODBCStmt[myNo]->stmt(),
                nullptr, nullptr, nullptr, nullptr,
                &coltype,
                nullptr, nullptr);
        }
        catch (RETCODE)
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
                                 odbc_raii_select::result_type const&   result,
                                 std::vector<SQLSMALLINT> const&        coltype,
                                 bool                                   rc = true)
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
                VARIANT elem = makeVariantFromSQLType(coltype[rc ? innerIter : outerIter], p);
                std::swap(*reinterpret_cast<VARIANT*>(innerBegin + innerIter * innerElemSize), elem);
                ::VariantClear(&elem);
            }
            reinterpret_cast<VARIANT*>(outerBegin + outerIter * outerElemSize)->vt = VT_ARRAY | VT_VARIANT;
            reinterpret_cast<VARIANT*>(outerBegin + outerIter * outerElemSize)->parray = innerArray.get();
        }
        ret.vt = VT_ARRAY | VT_VARIANT;
        ret.parray = outerArray.get();
    }

    // std::vector<VARIANT> ==> Variant()
    VARIANT vec2VArray(std::vector<VARIANT>& vec)
    {
        VARIANT ret;
        ::VariantInit(&ret);
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
        vec.clear();
        return ret;
    }

}