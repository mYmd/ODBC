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

    BSTR getBSTR(VARIANT* expr);

    class safearrayRAII {
        SAFEARRAY* pArray;
    public:
        safearrayRAII(SAFEARRAY* p) : pArray(p) {}
        ~safearrayRAII() { ::SafeArrayUnaccessData(pArray); }
        SAFEARRAY* get() const { return pArray; }
    };

    VARIANT vec2VArray(std::vector<VARIANT>&&);

    template <typename T, typename F>
    VARIANT vec2VArray(std::vector<T>&&, F&&);

    class header_getter    {
        std::vector<column_name_type>   v_colname;
        std::vector<SQLSMALLINT>        v_coltype;
    public:
        void operator()(std::vector<column_name_type>&  colname ,
                        std::vector<SQLSMALLINT>&               ,
                        std::vector<SQLULEN>&                   ,
                        std::vector<SQLSMALLINT>&               ,
                        std::vector<SQLSMALLINT>&       coltype ,
                        std::vector<SQLSMALLINT>&               ,
                        std::vector<SQLLEN>&        )
        {
            v_colname = colname;
            v_coltype = coltype;
        }
        void get(VARIANT& v)
        {
            auto bstr_trans = [](VARIANT& elem, column_name_type& c) {
                elem.vt = VT_BSTR;
                elem.bstrVal = SysAllocString(c.data());
            };
            auto type_trans = [](VARIANT& elem, SQLSMALLINT t) {
                elem.vt = VT_BSTR;
                tstring const str = getTypeStr(t);
                TCHAR const* p = str.empty() ? nullptr : &str[0];
                elem.bstrVal = SysAllocString(p);
            };
            VARIANT colname_array = vec2VArray(std::move(v_colname), bstr_trans);
            VARIANT coltype_array = vec2VArray(std::move(v_coltype), type_trans);
            std::vector<VARIANT> vec{colname_array, coltype_array};
            ::VariantClear(&v);
            std::swap(v, vec2VArray(std::move(vec)));
        }
    };

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

VARIANT __stdcall selectODBC_rowWise(__int32 myNo, VARIANT* SQL, VARIANT* header)
{
    VARIANT ret;
    ::VariantInit(&ret);
    BSTR bstr = getBSTR(SQL);
    if (!bstr || myNo < 0 || vODBCStmt.size() <= myNo)        return ret;
    std::vector<std::vector<VARIANT>> vec;
    std::vector<VARIANT> elem;
    auto init_func = [&](SQLSMALLINT c) {
        elem.resize(c);
    };
    auto elem_func = [&](SQLSMALLINT j, tstring const& str, SQLSMALLINT coltype) {
        TCHAR const* p = str.empty() ? 0 : &str[0];
        elem[j] = makeVariantFromSQLType(coltype, p);
    };
    auto add_func = [&](std::size_t) {
        vec.push_back(elem);
    };
    header_getter header_func;
    auto recordLen = select_table(  vODBCStmt[myNo]->stmt() ,
                                tstring{bstr}           ,
                            header_func             ,
                        init_func               ,
                    elem_func               ,
                add_func                );
    ::VariantClear(header);
    header_func.get(*header);
    std::vector<VARIANT> ret_vec;
    for (auto i = vec.begin(); i < vec.end(); ++i)
        ret_vec.push_back(vec2VArray(std::move(*i)));
    return vec2VArray(std::move(ret_vec));
}

VARIANT __stdcall selectODBC_columnWise(__int32 myNo, VARIANT* SQL, VARIANT* header)
{
    VARIANT ret;
    ::VariantInit(&ret);
    BSTR bstr = getBSTR(SQL);
    if (!bstr || myNo < 0 || vODBCStmt.size() <= myNo)        return ret;
    std::vector<std::vector<VARIANT>> vec;
    auto init_func = [&](SQLSMALLINT c) {
        vec.resize(c);
    };
    auto elem_func = [&](SQLSMALLINT j, tstring const& str, SQLSMALLINT coltype) {
        TCHAR const* p = str.empty() ? 0 : &str[0];
        vec[j].push_back(makeVariantFromSQLType(coltype, p));
    };
    auto add_func = [&](std::size_t) {    };
    header_getter header_func;
    auto recordLen = select_table(  vODBCStmt[myNo]->stmt() ,
                                tstring{bstr}           ,
                            header_func             ,
                        init_func               ,
                    elem_func               ,
                add_func                );
    ::VariantClear(header);
    header_func.get(*header);
    std::vector<VARIANT> ret_vec;
    for ( auto i = vec.begin(); i < vec.end(); ++i )
        ret_vec.push_back(vec2VArray(std::move(*i)));
    return vec2VArray(std::move(ret_vec));
}

VARIANT __stdcall selectODBC(__int32 myNo, VARIANT* SQL, VARIANT* header)
{
    VARIANT ret;
    ::VariantInit(&ret);
    BSTR bstr = getBSTR(SQL);
    if (!bstr || myNo < 0 || vODBCStmt.size() <= myNo)        return ret;
    std::vector<std::vector<VARIANT>> vec;
    std::vector<column_name_type>   v_colname;
    std::vector<SQLSMALLINT>        v_coltype;
    header_getter                   header_func;
    auto init_func = [&](SQLSMALLINT c) {
        vec.resize(c);
    };
    auto elem_func = [&](SQLSMALLINT j, tstring const& str, SQLSMALLINT coltype) {
        TCHAR const* p = str.empty() ? 0 : &str[0];
        vec[j].push_back(makeVariantFromSQLType(coltype, p));
    };
    auto add_func = [&](std::size_t) {};
    auto recordLen = select_table(  vODBCStmt[myNo]->stmt(),
                                tstring{bstr},
                            header_func,
                        init_func,
                    elem_func,
                add_func);
    //-----------
    ::VariantClear(header);
    header_func.get(*header);
    //---------------------------------
    SAFEARRAYBOUND rgb[2] = { { static_cast<ULONG>(recordLen), 0 }, { static_cast<ULONG>(vec.size()), 0 } };
    safearrayRAII pArray(::SafeArrayCreate(VT_VARIANT, 2, rgb));
    auto const elemsize = ::SafeArrayGetElemsize(pArray.get());
    char* it = nullptr;
    ::SafeArrayAccessData(pArray.get(), reinterpret_cast<void**>(&it));
    if (!it)
        return ret;
    for (std::size_t col = 0; col < vec.size(); ++col)
    {
        for (std::size_t row = 0; row < recordLen; ++row)
        {
            std::swap(*reinterpret_cast<VARIANT*>(it), vec[col][row]);
            it += elemsize;
        }
    }
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = pArray.get();
    return ret;
}

VARIANT __stdcall columnAttributes(__int32 myNo, VARIANT* SQL)
{
    VARIANT ret;    ::VariantInit(&ret);
    BSTR bstr = getBSTR(SQL);
    if ( !bstr || myNo < 0 || vODBCStmt.size() <= myNo )    return ret;
    //------------------------
    header_getter                   header_func;
    SQLSMALLINT     len = columnAttribute(  vODBCStmt[myNo]->stmt() ,
                                        tstring(bstr)           ,
                                    nullptr                 ,
                                header_func             ,
                            true                    );
    if ( len == 0 )     return ret;
    header_func.get(ret);
    return ret;
}

VARIANT __stdcall execODBC(__int32 myNo, VARIANT* SQLs)
{
    VARIANT ret;
    ::VariantInit(&ret);
    if ( myNo < 0 || vODBCStmt.size() <= myNo )         return ret;
    if (!SQLs ||  0 == (VT_ARRAY & SQLs->vt))           return ret;
    SAFEARRAY* pArray = (0 == (VT_BYREF & SQLs->vt)) ? (SQLs->parray) : (*SQLs->pparray);
    if (!pArray || 1 != ::SafeArrayGetDim(pArray))      return ret;
    SAFEARRAYBOUND bounds = { 1,0 };
    {
        ::SafeArrayGetLBound(pArray, 1, &bounds.lLbound);
        LONG ub = 0;
        ::SafeArrayGetUBound(pArray, 1, &ub);
        bounds.cElements = 1 + ub - bounds.lLbound;
    }
    cursor_colser       c_closer(vODBCStmt[myNo]->stmt(), true);
    VARIANT elem;
    ::VariantInit(&elem);
    std::vector<LONG> errorNo;
    for (ULONG i = 0; i < bounds.cElements; ++i)
    {
        LONG index = static_cast<LONG>(i) + bounds.lLbound;
        ::SafeArrayGetElement(pArray, &index, &elem);
        if (elem.vt == VT_BSTR && elem.bstrVal)
        {
            auto const rc = execDirect(tstring(elem.bstrVal), vODBCStmt[myNo]->stmt());
            if (rc != SQL_SUCCESS && rc != SQL_SUCCESS_WITH_INFO)
                errorNo.push_back(index);
        }
        ::VariantClear(&elem);
    }
    auto errorNo_trans = [](VARIANT& elem, LONG c) {
        elem.vt = VT_I4;
        elem.lVal = c;
    };
    if (errorNo.size())
        ret = vec2VArray(std::move(errorNo), errorNo_trans);
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
    TCHAR const zeroStr [] {_T("")};
    auto push_back_func = [&](TCHAR const* p) {
        vec.push_back(makeVariantFromSQLType(SQL_CHAR, p? p: zeroStr));
    };
    catalogValue(table_func, st, 2, push_back_func);    //TABLE_SCHEM
    VARIANT schem_name = vec2VArray(std::move(vec));
    catalogValue(table_func, st, 3, push_back_func);    //TABLE_NAME
    VARIANT table_name = vec2VArray(std::move(vec));
    catalogValue(table_func, st, 4, push_back_func);    //TABLE_TYPE
    VARIANT type_name = vec2VArray(std::move(vec));
    vec.push_back(schem_name);
    vec.push_back(table_name);
    vec.push_back(type_name);
    return vec2VArray(std::move(vec));
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
    auto value_type = SQL_CHAR;
    TCHAR const zeroStr[]{ _T("") };
    auto push_back_func = [&](TCHAR const* p) {
        vec.push_back(makeVariantFromSQLType(SQL_CHAR, p ? p : zeroStr));
    };

    catalogValue(primarykeys_func, st, 4, push_back_func);  // KEY_NAME
    VARIANT primarykeys = vec2VArray(std::move(vec));

    catalogValue(column_func, st, 4, push_back_func);       // COLUMN_NAME
    VARIANT column_name = vec2VArray(std::move(vec));

    catalogValue(column_func, st, 6, push_back_func);       // TYPE_NAME
    VARIANT type_name = vec2VArray(std::move(vec));

    value_type = SQL_SMALLINT;
    catalogValue(column_func, st, 11, push_back_func);      // IS_NULLABLE(SQL_SMALLINT)
    VARIANT is_nullable = vec2VArray(std::move(vec));

    value_type = SQL_INTEGER;
    catalogValue(column_func, st, 7, push_back_func);       // COLUMN_SIZE(SQL_INTEGER)
    VARIANT column_size = vec2VArray(std::move(vec));

    catalogValue(column_func, st, 17, push_back_func);      // ORDINAL_POSITION(SQL_INTEGER)
    VARIANT ordinal_position = vec2VArray(std::move(vec));

    vec.push_back(column_name);
    vec.push_back(type_name);
    vec.push_back(column_size);
    vec.push_back(is_nullable);
    vec.push_back(ordinal_position);
    vec.push_back(primarykeys);

    return vec2VArray(std::move(vec));
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
            ret.bstrVal = ::SysAllocString(strln);
            return ret;
        }
        case SQL_SMALLINT:
        case SQL_INTEGER:
        case SQL_BIT:
        case SQL_TINYINT:
        {
            long lOut;
            auto const vdr = ::VarI4FromStr(strln, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &lOut);
            ret.vt = VT_I4;
            ret.lVal = lOut;
            return ret;
        }
        case SQL_BIGINT:
        {
            LONG64  i64Out;
            auto const vdr = ::VarI8FromStr(strln, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &i64Out);
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
            auto const vdr = ::VarR8FromStr(strln, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &dOut);
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
            auto const vdr = ::VarDateFromStr(strln2, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &dOut);
            ret.vt = VT_DATE;
            ret.date = dOut;
            return ret;
        }
        default:
            ret.vt = VT_NULL;
            return ret;
        }
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

    // std::vector<VARIANT> ==> Variant()
    VARIANT vec2VArray(std::vector<VARIANT>&& vec)
    {
        VARIANT ret;
        ::VariantInit(&ret);
        SAFEARRAYBOUND rgb = { static_cast<ULONG>(vec.size()), 0 };
        safearrayRAII pArray(::SafeArrayCreate(VT_VARIANT, 1, &rgb));
        char* it = nullptr;
        ::SafeArrayAccessData(pArray.get(), reinterpret_cast<void**>(&it));
        if (!it)            return ret;
        auto const elemsize = ::SafeArrayGetElemsize(pArray.get());
        for (std::size_t i = 0; i < vec.size(); ++i)
            std::swap(*reinterpret_cast<VARIANT*>(it + i * elemsize), vec[i]);
        ret.vt = VT_ARRAY | VT_VARIANT;
        ret.parray = pArray.get();
        vec.clear();
        return ret;
    }

    // std::vector<T> ==> Variant()
    template <typename T, typename F>
    VARIANT vec2VArray(std::vector<T>&& vec, F&& trans)
    {
        VARIANT ret;
        ::VariantInit(&ret);
        SAFEARRAYBOUND rgb = { static_cast<ULONG>(vec.size()), 0 };
        safearrayRAII pArray(::SafeArrayCreate(VT_VARIANT, 1, &rgb));
        char* it = nullptr;
        ::SafeArrayAccessData(pArray.get(), reinterpret_cast<void**>(&it));
        if (!it)            return ret;
        auto const elemsize = ::SafeArrayGetElemsize(pArray.get());
        for (std::size_t i = 0; i < vec.size(); ++i)
            std::forward<F>(trans)(*reinterpret_cast<VARIANT*>(it + i * elemsize), vec[i]);
        ret.vt = VT_ARRAY | VT_VARIANT;
        ret.parray = pArray.get();
        vec.clear();
        return ret;
    }

}
