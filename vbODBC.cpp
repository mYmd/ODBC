//vbODBC.cpp
//Copyright (c) 2015 mmYYmmdd
#include "stdafx.h"
#include "odbcResource.hpp"
#include <memory>
#include <OleAuto.h>

using namespace mymd;

namespace {

    std::vector<std::unique_ptr<odbc_set>>  vODBCStmt;

    __int32 vODBCStmt_size()
    {
        return static_cast<__int32>(vODBCStmt.size());
    }

    VARIANT makeVariantFromSQLType(SQLSMALLINT, LPCOLESTR);

    BSTR getBSTR(VARIANT* expr);

    class safearrayRAII {
        SAFEARRAY* pArray;
    public:
        safearrayRAII(SAFEARRAY* p) : pArray(p) {}
        ~safearrayRAII() { ::SafeArrayUnaccessData(pArray); }
        SAFEARRAY* get() const { return pArray; }
    };

    VARIANT iVariant(VARTYPE t = VT_EMPTY)
    {
        VARIANT ret;
        ::VariantInit(&ret);
        ret.vt = t;
        return ret;
    }

    template <typename Container_t>
    VARIANT vec2VArray(Container_t&&);

    template <typename Container_t, typename F>
    VARIANT vec2VArray(Container_t&&, F&&);

    // 
    class header_getter    {
        std::vector<column_t::name_type>    v_colname;
        std::vector<SQLSMALLINT>            v_coltype;
    public:
        void operator()(std::vector<column_t::name_type>&   colname ,
                        std::vector<SQLSMALLINT>&                   ,
                        std::vector<SQLULEN>&                       ,
                        std::vector<SQLSMALLINT>&                   ,
                        std::vector<SQLSMALLINT>&           coltype ,
                        std::vector<SQLSMALLINT>&                   )
        {
            v_colname = std::move(colname);
            v_coltype = std::move(coltype);
        }
        VARIANT getHeader();
    };

    std::vector<std::vector<VARIANT>>
        selectODBC_columnWise_imple(odbc_raii_statement&, BSTR, VARIANT*);

}

//----------------------------------------------------------------------

void __stdcall terminateODBC(__int32 myNo)
{
    if ( 0 <= myNo && myNo < vODBCStmt_size() )
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
        auto p = std::make_unique<odbc_set>(connectName);
        if ( p->isError() )
        {
            ::VariantClear(rawStr);
            *rawStr = makeVariantFromSQLType(SQL_CHAR, p->errorMessage().data());
            return -1;
        }
        if ( 0 <= myNo && myNo < vODBCStmt_size() )
        {
            vODBCStmt[myNo] = std::move(p);
        }
        else
        {
            vODBCStmt.push_back(std::move(p));
            myNo = static_cast<int>(vODBCStmt.size() - 1);
        }
    }
    catch (...)
    {
        ::VariantClear(rawStr);
        *rawStr = makeVariantFromSQLType(SQL_CHAR, _T("UnKnown Error"));
        return -1;
    }
    return myNo;
}

__int32 __stdcall setQueryTimeout(__int32 myNo, __int32 sec)
{
    if (myNo < 0 || vODBCStmt_size() <= myNo)
        return 0;
    if (sec < 0)    sec = 0;
    auto val = static_cast<SQLULEN>(sec);
    auto p = static_cast<SQLPOINTER>(&val);
    auto setSTMT = [=](HSTMT x) {
        return ::SQLSetStmtAttr(x, SQL_ATTR_QUERY_TIMEOUT, p, SQL_IS_POINTER);//0
    };
    auto result = vODBCStmt[myNo]->stmt().invoke(setSTMT);
    return (result == SQL_SUCCESS || result == SQL_SUCCESS_WITH_INFO) ? sec : 0;
}

VARIANT __stdcall getStatementError(__int32 myNo)
{
    if (myNo < 0 || vODBCStmt_size() <= myNo)
        return iVariant();
    SQLDiagRec<>    diagRec;
    vODBCStmt[myNo]->stmt().invoke(diagRec);
    return makeVariantFromSQLType(SQL_CHAR, diagRec.getMessage().data());
}

VARIANT __stdcall selectODBC_rowWise(__int32 myNo, VARIANT* SQL, VARIANT* header)
{
    BSTR bstr = getBSTR(SQL);
    if (!bstr || myNo < 0 || vODBCStmt_size() <= myNo)        return iVariant();
    std::vector<VARIANT> vec;
    std::vector<VARIANT> elem;
    SQLSMALLINT col_N{0};
    auto init_func = [&](SQLSMALLINT c) {
        elem.resize(col_N = c);
    };
    auto elem_func = [&](SQLSMALLINT j, TCHAR const* str, SQLSMALLINT coltype) {
        elem[j] = makeVariantFromSQLType(coltype, str);
    };
    auto add_func = [&](std::size_t) {
        vec.push_back(vec2VArray(std::move(elem)));
        elem.resize(col_N);
    };
    header_getter header_func;
    auto recordLen = select_table(  vODBCStmt[myNo]->stmt() ,
                                tstring{bstr}           ,
                            header_func             ,
                        init_func               ,
                    elem_func               ,
                add_func                );
    ::VariantClear(header);
    std::swap(*header, header_func.getHeader());
    return vec2VArray(std::move(vec));
}

VARIANT __stdcall selectODBC_columnWise(__int32 myNo, VARIANT* SQL_expr, VARIANT* header)
{
    BSTR sql = getBSTR(SQL_expr);
    if (!sql || myNo < 0 || vODBCStmt_size() <= myNo)       return iVariant();
    auto vec = selectODBC_columnWise_imple(vODBCStmt[myNo]->stmt(), sql, header);
    std::vector<VARIANT> ret_vec;
    for (auto i = vec.begin(); i < vec.end(); ++i)
        ret_vec.push_back(vec2VArray(std::move(*i)));
    return vec2VArray(std::move(ret_vec));
}

VARIANT __stdcall selectODBC(__int32 myNo, VARIANT* SQL_expr, VARIANT* header)
{
    BSTR sql = getBSTR(SQL_expr);
    if (!sql || myNo < 0 || vODBCStmt_size() <= myNo)       return iVariant();
    auto vec = selectODBC_columnWise_imple(vODBCStmt[myNo]->stmt(), sql, header);
    if ( !vec.size() )      return iVariant();
    auto recordLen = vec[0].size();
    SAFEARRAYBOUND rgb[2] = { { static_cast<ULONG>(recordLen), 0 }, { static_cast<ULONG>(vec.size()), 0 } };
    safearrayRAII pArray(::SafeArrayCreate(VT_VARIANT, 2, rgb));
    auto const elemsize = ::SafeArrayGetElemsize(pArray.get());
    char* it{nullptr};
    ::SafeArrayAccessData(pArray.get(), reinterpret_cast<void**>(&it));
    if (!it)        return iVariant();
    for (std::size_t col = 0; col < vec.size(); ++col)
    {
        for (std::size_t row = 0; row < recordLen; ++row)
        {
            std::swap(*reinterpret_cast<VARIANT*>(it), vec[col][row]);
            it += elemsize;
        }
    }
    VARIANT ret = iVariant(VT_ARRAY | VT_VARIANT);
    ret.parray = pArray.get();
    return ret;
}

VARIANT __stdcall columnAttributes(__int32 myNo, VARIANT* SQL)
{
    BSTR bstr = getBSTR(SQL);
    if ( !bstr || myNo < 0 || vODBCStmt_size() <= myNo )    return iVariant();
    //------------------------
    header_getter                   header_func;
    SQLSMALLINT     len = columnAttribute(  vODBCStmt[myNo]->stmt() ,
                                        tstring(bstr)           ,
                                    nullptr,
                                nullptr,
                            header_func             ,
                        true                    );
    if ( len == 0 )     return iVariant();
    return header_func.getHeader();
}

VARIANT __stdcall execODBC(__int32 myNo, VARIANT* SQLs)
{
    if ( myNo < 0 || vODBCStmt_size() <= myNo )         return iVariant();
    if (!SQLs ||  0 == (VT_ARRAY & SQLs->vt))           return iVariant();
    SAFEARRAY* pArray = (0 == (VT_BYREF & SQLs->vt)) ? (SQLs->parray) : (*SQLs->pparray);
    if (!pArray || 1 != ::SafeArrayGetDim(pArray))      return iVariant();
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
    std::vector<LONG>       errorNo;
    std::vector<VARIANT>    errorMessaged;
    SQLDiagRec<>            diagRec;
    for (ULONG i = 0; i < bounds.cElements; ++i)
    {
        LONG index = static_cast<LONG>(i) + bounds.lLbound;
        ::SafeArrayGetElement(pArray, &index, &elem);
        if (elem.vt == VT_BSTR && elem.bstrVal)
        {
            auto const rc = execDirect(tstring(elem.bstrVal), vODBCStmt[myNo]->stmt());
            if (rc != SQL_SUCCESS && rc != SQL_SUCCESS_WITH_INFO)
            {
                errorNo.push_back(index);
                vODBCStmt[myNo]->stmt().invoke(diagRec);
                errorMessaged.push_back(makeVariantFromSQLType(SQL_CHAR, diagRec.getMessage().data()));
            }
        }
        ::VariantClear(&elem);
    }
    auto errorNo_trans = [](LONG c) {
        VARIANT elem = iVariant(VT_I4);
        elem.lVal = c;
        return elem;
    };
    return (errorNo.size())?
        vec2VArray(std::vector<VARIANT>{vec2VArray(std::move(errorNo), errorNo_trans), 
                                        vec2VArray(std::move(errorMessaged))          } )
        : iVariant();
}

// テーブル一覧
VARIANT __stdcall table_list_all(__int32 myNo, VARIANT* schemaName)
{
    BSTR schema_name_b = getBSTR(schemaName);
    if (!schema_name_b || myNo < 0 || vODBCStmt_size() <= myNo)
        return iVariant();;
    tstring schema_name_t(schema_name_b);
    SQLTCHAR* schema_name = const_cast<SQLTCHAR*>(static_cast<const SQLTCHAR*>(schema_name_t.c_str()));
    auto schema_len = static_cast<SQLSMALLINT>(schema_name_t.length());
    if (schema_len == 0)      schema_name = NULL;
    auto table_func = [=](HSTMT x)  {
        return ::SQLTables(x, NULL, SQL_NTS, schema_name, schema_len, NULL, SQL_NTS, NULL, SQL_NTS);
    };
    std::array<SQLUSMALLINT, 7> columns = {2, 3, 4};
    auto scheme_name_type = catalogValue(table_func             , 
                                         vODBCStmt[myNo]->stmt(),
                                         columns.begin()        ,
                                         columns.end()          );
    auto trans = [](tstring& s) {   return makeVariantFromSQLType(SQL_CHAR, &s[0]); };
    std::vector<VARIANT> vec;
    vec.push_back(vec2VArray(std::move(scheme_name_type[0]), trans));
    vec.push_back(vec2VArray(std::move(scheme_name_type[1]), trans));
    vec.push_back(vec2VArray(std::move(scheme_name_type[2]), trans));
    return vec2VArray(std::move(vec));
}

// https://www.ibm.com/support/knowledgecenter/ja/SSEPEK_11.0.0/odbc/src/tpc/db2z_fncolumns.html
//************************************************
//  1   TABLE_CAT           VARCHAR(128)
//  2   TABLE_SCHEM         VARCHAR(128)
//  3   TABLE_NAME          VARCHAR(128) NOT NULL
//  4   COLUMN_NAME         VARCHAR(128) NOT NULL
//  5   DATA_TYPE           SMALLINT NOT NULL
//  6   TYPE_NAME           VARCHAR(128) NOT NULL
//  7   COLUMN_SIZE         INTEGER
//  8   BUFFER_LENGTH       INTEGER
//  9   DECIMAL_DIGITS      SMALLINT
// 10   NUM_PREC_RADIX      SMALLINT
// 11   NULLABLE            SMALLINT NOT NULL
// 12   REMARKS             VARCHAR(762)
// 13   COLUMN_DEF          VARCHAR(254)
// 14   SQL_DATA_TYPE       SMALLINT NOT NULL
// 15   SQL_DATETIME_SUB    SMALLINT
// 16   CHAR_OCTET_LENGTH   INTEGER
// 17   ORDINAL_POSITION    INTEGER NOT NULL
// 18   IS_NULLABLE         VARCHAR(254)
//************************************************

// テーブルにある全カラムの属性
VARIANT __stdcall columnAttributes_all(__int32 myNo, VARIANT* schemaName, VARIANT* tableName)
{
    BSTR schema_name_b{getBSTR(schemaName)}, table_Name_b{getBSTR(tableName)};
    if (!schema_name_b || !table_Name_b || myNo < 0 || vODBCStmt_size() <= myNo)
        return iVariant();
    tstring schema_name_t(schema_name_b), table_name_t(table_Name_b);
    SQLTCHAR* schema_name = const_cast<SQLTCHAR*>(static_cast<const SQLTCHAR*>(schema_name_t.c_str()));
    SQLTCHAR* table_Name  = const_cast<SQLTCHAR*>(static_cast<const SQLTCHAR*>(table_name_t.c_str()));
    auto schema_len = static_cast<SQLSMALLINT>(schema_name_t.length());
    auto table_len = static_cast<SQLSMALLINT>(table_name_t.length());
    if (schema_len == 0)      schema_name = NULL;
    auto const& st = vODBCStmt[myNo]->stmt();
    struct trans {  // Workaround for VC++2013
        SQLSMALLINT x;
        trans(SQLSMALLINT i) : x(i) {   }
        VARIANT operator ()(tstring& s) const
        {   return makeVariantFromSQLType(x, &s[0]);    };
    };
    //auto trans = [](SQLSMALLINT x) {
    //    return [=](tstring& s)  {   return makeVariantFromSQLType(x, &s[0]);    };
    //};
    std::vector<VARIANT> vec;
    {
        auto column_func = [=](HSTMT x) {
            return ::SQLColumns(x, NULL, SQL_NTS, schema_name, schema_len, table_Name, table_len, NULL, SQL_NTS);
        };
        std::array<SQLUSMALLINT, 7> columns = { 4, 6, 11, 9, 5, 7, 17 };
        auto column_attr = catalogValue(column_func, st, columns.begin(), columns.end());
        VARIANT column_name = vec2VArray(std::move(column_attr[0]), trans(SQL_CHAR));
        VARIANT type_name = vec2VArray(std::move(column_attr[1]), trans(SQL_CHAR));
        VARIANT is_nullable = vec2VArray(std::move(column_attr[2]), trans(SQL_SMALLINT));
        VARIANT Decimal_Digits = vec2VArray(std::move(column_attr[3]), trans(SQL_SMALLINT));
        VARIANT column_size = vec2VArray(std::move(column_attr[5]), trans(SQL_INTEGER));
        VARIANT ordinal_position = vec2VArray(std::move(column_attr[6]), trans(SQL_INTEGER));
        vec.push_back(column_name);         // 0
        vec.push_back(type_name);           // 1
        vec.push_back(column_size);         // 2
        vec.push_back(Decimal_Digits);      // 3
        vec.push_back(is_nullable);         // 4
        vec.push_back(ordinal_position);    // 5
    }
    {
        auto primarykeys_func = [=](HSTMT x) {
            return ::SQLPrimaryKeys(x, NULL, SQL_NTS, schema_name, schema_len, table_Name, table_len);
        };
        SQLUSMALLINT keycolumns[] = { 4 };
        auto key_value = catalogValue(primarykeys_func, st, keycolumns, keycolumns + 1);    //KEY_NAME
        VARIANT primarykeys = vec2VArray(std::move(key_value[0]), trans(SQL_CHAR));
        vec.push_back(primarykeys);         // 6
    }
    return vec2VArray(std::move(vec));
}

namespace {

    VARIANT makeVariantFromSQLType(SQLSMALLINT type, LPCOLESTR expr)
    {
        VARIANT ret;    ::VariantInit(&ret);    ret.vt = VT_NULL;
        if (!expr)      return ret;
        switch (type)
        {
        case SQL_CHAR:      case SQL_VARCHAR:       case SQL_LONGVARCHAR:
        case SQL_WCHAR:     case SQL_WVARCHAR:      case SQL_WLONGVARCHAR:
        case SQL_BINARY:    case SQL_VARBINARY:     case SQL_LONGVARBINARY:
        {
            ret.vt = VT_BSTR;
            ret.bstrVal = ::SysAllocString(expr);
            return ret;
        }
        case SQL_SMALLINT:  case SQL_INTEGER:   case SQL_BIT:   case SQL_TINYINT:
        {
            long lOut;
            auto const vdr = ::VarI4FromStr(expr, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &lOut);
            ret.vt = VT_I4;
            ret.lVal = lOut;
            return ret;
        }
        case SQL_BIGINT:
        {
            LONG64  i64Out;
            auto const vdr = ::VarI8FromStr(expr, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &i64Out);
            ret.vt = VT_I8;
            ret.llVal = i64Out;
            return ret;
        }
        case SQL_NUMERIC:   case SQL_DECIMAL:   case SQL_FLOAT: case SQL_REAL:  case SQL_DOUBLE:
        {
            double dOut;
            auto const vdr = ::VarR8FromStr(expr, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &dOut);
            ret.vt = VT_R8;
            ret.dblVal = dOut;
            return ret;
        }
        case SQL_TYPE_DATE: case SQL_TYPE_TIME: case SQL_TYPE_TIMESTAMP:
        {
            OLECHAR date_expr[] = _T("2001-01-01 00:00:00");
            auto p = expr;
            auto q = date_expr;
            while (*p != _T('\0') && *p != _T('.') && *q != _T('\0'))       *q++ = *p++;
            *q = _T('\0');
            DATE dOut;
            auto const vdr = ::VarDateFromStr(date_expr, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &dOut);
            ret.vt = VT_DATE;
            ret.date = dOut;
            return ret;
        }
        default:    return ret;
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
    template <typename Container_t>
    VARIANT vec2VArray(Container_t&& cont)
    {
        static_assert(!std::is_reference<Container_t>::value, "vec2VArray's parameter is a rvalue reference !!");
        auto trans = [](typename Container_t::reference x) -> typename Container_t::reference   {
            return x;
        };
        return vec2VArray(std::move(cont), trans);
    }

    // std::vector<T> ==> Variant()
    template <typename Container_t, typename F>
    VARIANT vec2VArray(Container_t&& cont, F&& trans)
    {
        static_assert(!std::is_reference<Container_t>::value, "vec2VArray's parameter is a rvalue reference !!");
        SAFEARRAYBOUND rgb = { static_cast<ULONG>(cont.size()), 0 };
        safearrayRAII pArray(::SafeArrayCreate(VT_VARIANT, 1, &rgb));
        char* it = nullptr;
        ::SafeArrayAccessData(pArray.get(), reinterpret_cast<void**>(&it));
        if (!it)            return iVariant();
        auto const elemsize = ::SafeArrayGetElemsize(pArray.get());
        std::size_t i{0};
        for (auto p = cont.begin(); p != cont.end(); ++p, ++i)
            std::swap(*reinterpret_cast<VARIANT*>(it + i * elemsize), std::forward<F>(trans)(*p));
        VARIANT ret = iVariant(VT_ARRAY | VT_VARIANT);
        ret.parray = pArray.get();
        cont.clear();
        return ret;
    }

    std::vector<std::vector<VARIANT>>
        selectODBC_columnWise_imple(odbc_raii_statement& st, BSTR sql, VARIANT* header)
    {
        std::vector<std::vector<VARIANT>> vec;
        auto init_func = [&](SQLSMALLINT c) {
            vec.resize(c);
        };
        auto elem_func = [&](SQLSMALLINT j, TCHAR const* str, SQLSMALLINT coltype) {
            vec[j].push_back(makeVariantFromSQLType(coltype, str));
        };
        auto add_func = [&](std::size_t) {};
        header_getter header_func;
        auto recordLen = select_table(  st,
                                    tstring{ sql },
                                header_func,
                            init_func,
                        elem_func,
                    add_func);
        ::VariantClear(header);
        std::swap(*header, header_func.getHeader());
        return vec;
    }

    VARIANT header_getter::getHeader()
    {
        auto bstr_trans = [](column_t::name_type& c) {
            VARIANT elem = iVariant(VT_BSTR);
            elem.bstrVal = SysAllocString(c.data());
            return elem;
        };
        auto type_trans = [](SQLSMALLINT t) {
            VARIANT elem = iVariant(VT_BSTR);
            tstring const str = getTypeStr(t);
            TCHAR const* p = str.empty() ? nullptr : &str[0];
            elem.bstrVal = SysAllocString(p);
            return elem;
        };
        VARIANT colname_array = vec2VArray(std::move(v_colname), bstr_trans);
        VARIANT coltype_array = vec2VArray(std::move(v_coltype), type_trans);
        std::vector<VARIANT> vec{ colname_array, coltype_array };
        return vec2VArray(std::move(vec));
    }

}
