//vbODBC.cpp
//Copyright (c) 2015 mmYYmmdd
#include "stdafx.h"
#include "odbcResource.hpp"
#include <memory>
#include <OleAuto.h>
#include <fstream>
#include <codecvt>

using namespace mymd;

namespace {

    std::vector<std::unique_ptr<odbc_set>>  vODBCStmt;

    __int32 vODBCStmt_size() noexcept
    {
        return static_cast<__int32>(vODBCStmt.size());
    }

    VARIANT makeVariantFromSQLType(SQLSMALLINT, LPCOLESTR) noexcept;

    BSTR getBSTR(VARIANT const&) noexcept;
    BSTR getBSTR(VARIANT const*) noexcept;

    struct SafeArrayUnaccessor {
        void operator()(SAFEARRAY* ptr) const  noexcept
        { ::SafeArrayUnaccessData(ptr); }
    };

    using safearrayRAII = std::unique_ptr<SAFEARRAY, SafeArrayUnaccessor>;

    VARIANT iVariant(VARTYPE t = VT_EMPTY) noexcept
    {
        VARIANT ret;
        ::VariantInit(&ret);
        ret.vt = t;
        return ret;
    }

    template <typename Container_t>
    VARIANT vec2VArray(Container_t&&) noexcept;

    template <typename Container_t, typename F>
    VARIANT vec2VArray(Container_t&&, F&&) noexcept;

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
                        std::vector<SQLSMALLINT>&                   ) noexcept
        {
            v_colname = std::move(colname);
            v_coltype = std::move(coltype);
        }
        VARIANT getHeader() noexcept;
    };

    std::vector<std::vector<VARIANT>>
        selectODBC_columnWise_imple(odbc_raii_statement&, BSTR, VARIANT&);

    class fputws_t  {
        struct file_closer {
            void operator ()(FILE* pF) const  noexcept { if (pF) ::fclose(pF); }
        };
        using fileCloseRAII = std::unique_ptr<FILE, file_closer>;
        fileCloseRAII   fc_RAII;
        __int32 codepage_;
        bool CrLf_;
    public:
        fputws_t(BSTR fileName, __int32 Code_Page, bool append, bool CrLf) noexcept;
        void lineFeed() const noexcept;
        bool put(std::wstring const& x) const noexcept;
        void flush() const noexcept;
    };

}

//----------------------------------------------------------------------
void __stdcall terminateODBC(__int32 myNo) noexcept
{
    if ( 0 <= myNo && myNo < vODBCStmt_size() )
    {
        try {   vODBCStmt[myNo].reset();    }
        catch (const std::exception&)   { }
    }
}

void __stdcall terminateODBC_all() noexcept
{
    try { vODBCStmt.clear();    }
    catch (const std::exception&) { }
}

__int32 __stdcall initODBC(__int32& myNo, VARIANT& rawStr) noexcept
{
    auto bstr = getBSTR(rawStr);
    if (!bstr)                     return -1;
    try
    {
        std::wstring connectName{ bstr };
        auto p = std::make_unique<odbc_set>(connectName);
        if ( p->isError() )
        {
            ::VariantClear(&rawStr);
            rawStr = makeVariantFromSQLType(SQL_CHAR, p->errorMessage().data());
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
    catch (const std::exception&)
    {
        ::VariantClear(&rawStr);
        rawStr = makeVariantFromSQLType(SQL_CHAR, L"UnKnown Error");
        return -1;
    }
    return myNo;
}

__int32 __stdcall setQueryTimeout(__int32 myNo, __int32 sec) noexcept
{
    if (myNo < 0 || vODBCStmt_size() <= myNo)
        return 0;
    if (sec < 0)    sec = 0;
    auto val = static_cast<SQLULEN>(sec);
    auto p = static_cast<SQLPOINTER>(&val);
    auto setSTMT = [=](HSTMT x) {
        return ::SQLSetStmtAttr(x, SQL_ATTR_QUERY_TIMEOUT, p, SQL_IS_POINTER);//0
    };
    try
    {
        auto result = vODBCStmt[myNo]->stmt().invoke(setSTMT);
        return (result == SQL_SUCCESS || result == SQL_SUCCESS_WITH_INFO) ? sec : 0;
    }
    catch (const std::exception&)
    {
        return 0;
    }
}

VARIANT __stdcall getStatementError(__int32 myNo) noexcept
{
    if (myNo < 0 || vODBCStmt_size() <= myNo)
        return iVariant();
    try
    {
        SQLDiagRec<> diagRec;
        vODBCStmt[myNo]->stmt().invoke(diagRec);
        return makeVariantFromSQLType(SQL_CHAR, diagRec.getMessage().data());
    }
    catch (const std::exception&)
    {
        return iVariant();
    }
}

VARIANT __stdcall selectODBC_rowWise(__int32 myNo, VARIANT const& SQL, VARIANT& header) noexcept
{
    auto bstr = getBSTR(SQL);
    if (!bstr || myNo < 0 || vODBCStmt_size() <= myNo)        return iVariant();
    try
    {
        std::vector<VARIANT> vec;
        std::vector<VARIANT> elem;
        SQLSMALLINT col_N{ 0 };
        auto init_func = [&](SQLSMALLINT c) {
            elem.resize(col_N = c);
        };
        auto elem_func = [&](SQLSMALLINT j, wchar_t const* str, SQLSMALLINT coltype) {
            elem[j] = makeVariantFromSQLType(coltype, str);
        };
        auto add_func = [&](std::size_t) {
            vec.push_back(vec2VArray(std::move(elem)));
            elem.resize(col_N);
        };
        header_getter   header_func;
        auto recordLen = select_table(  vODBCStmt[myNo]->stmt(),
                                        std::wstring{ bstr },
                                        header_func,
                                        init_func,
                                        elem_func,
                                        add_func);
        ::VariantClear(&header);
        std::swap(header, header_func.getHeader());
        return vec2VArray(std::move(vec));
    }
    catch (const std::exception&)
    {
        return iVariant();
    }
}

VARIANT __stdcall selectODBC_columnWise(__int32 myNo, VARIANT const& SQL_expr, VARIANT& header) noexcept
{
    auto sql = getBSTR(SQL_expr);
    if (!sql || myNo < 0 || vODBCStmt_size() <= myNo)       return iVariant();
    try
    {
        auto vec = selectODBC_columnWise_imple(vODBCStmt[myNo]->stmt(), sql, header);
        std::vector<VARIANT> ret_vec;
        ret_vec.reserve(vec.size());
        for (auto i = vec.begin(); i < vec.end(); ++i)
            ret_vec.push_back(vec2VArray(std::move(*i)));
        return vec2VArray(std::move(ret_vec));
    }
    catch (const std::exception&)
    {
        return iVariant();
    }
}

VARIANT __stdcall selectODBC(__int32 myNo, VARIANT const& SQL_expr, VARIANT& header) noexcept
{
    auto sql = getBSTR(SQL_expr);
    if (!sql || myNo < 0 || vODBCStmt_size() <= myNo)       return iVariant();
    try
    {
        auto vec = selectODBC_columnWise_imple(vODBCStmt[myNo]->stmt(), sql, header);
        if (!vec.size())      return iVariant();
        auto recordLen = vec[0].size();
        SAFEARRAYBOUND rgb[2] = { { static_cast<ULONG>(recordLen), 0 },{ static_cast<ULONG>(vec.size()), 0 } };
        safearrayRAII pArray{ ::SafeArrayCreate(VT_VARIANT, 2, rgb) };
        auto const elemsize = ::SafeArrayGetElemsize(pArray.get());
        char* it{ nullptr };
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
        auto ret = iVariant(VT_ARRAY | VT_VARIANT);
        ret.parray = pArray.get();
        return ret;
    }
    catch (const std::exception&)
    {
        return iVariant();
    }
}

#if  defined _M_X64
using vbLongLong = __int64;
#else
using vbLongLong = __int32;
#endif

// 行ごとのコールバック
__int32 __stdcall selectODBC_map(__int32 myNo, VARIANT const& SQL, VARIANT& header, vbLongLong fun, VARIANT& param) noexcept
{
    using vbCallbackFunc_t = __int32 (__stdcall *)(__int32, VARIANT&, VARIANT&);
    auto callback = reinterpret_cast<vbCallbackFunc_t>(fun);
    auto bstr = getBSTR(SQL);
    if (!callback || !bstr || myNo < 0 || vODBCStmt_size() <= myNo)        return 0;
    try
    {
        std::vector<VARIANT> elem;
        SQLSMALLINT col_N{ 0 };
        __int32 ret{ 0 };
        auto init_func = [&](SQLSMALLINT c) {
            elem.resize(col_N = c);
        };
        auto elem_func = [&](SQLSMALLINT j, wchar_t const* str, SQLSMALLINT coltype) {
            elem[j] = makeVariantFromSQLType(coltype, str);
        };
        auto add_func = [&](std::size_t n) {
            auto v = vec2VArray(std::move(elem));
            auto c = callback(ret = static_cast<__int32>(n), v, param);
            ++ret;
            ::VariantClear(&v);
            elem.resize(col_N);
            return 0 != c;      // ここ
        };
        header_getter   header_func;
        auto recordLen = select_table(  vODBCStmt[myNo]->stmt(),
                                      std::wstring{ bstr },
                                      header_func,
                                      init_func,
                                      elem_func,
                                      add_func);
        ::VariantClear(&header);
        std::swap(header, header_func.getHeader());
        return ret;
    }
    catch (const std::exception&)
    {
        return 0;
    }
}

VARIANT __stdcall columnAttributes(__int32 myNo, VARIANT const& SQL) noexcept
{
    auto bstr = getBSTR(SQL);
    if ( !bstr || myNo < 0 || vODBCStmt_size() <= myNo )    return iVariant();
    //------------------------
    header_getter   header_func;
    auto len = columnAttribute( vODBCStmt[myNo]->stmt(),
                                    std::wstring(bstr),
                                        nullptr,
                                            nullptr,
                                                header_func,
                                                    true);
    if ( len == 0 )     return iVariant();
    return header_func.getHeader();
}

VARIANT __stdcall execODBC(__int32 myNo, VARIANT const& SQLs) noexcept
{
    if ( myNo < 0 || vODBCStmt_size() <= myNo )         return iVariant();
    if (0 == (VT_ARRAY & SQLs.vt))                      return iVariant();
    auto pArray = (0 == (VT_BYREF & SQLs.vt)) ? (SQLs.parray) : (*SQLs.pparray);
    if (!pArray || 1 != ::SafeArrayGetDim(pArray))      return iVariant();
    SAFEARRAYBOUND bounds = { 1,0 };
    {
        ::SafeArrayGetLBound(pArray, 1, &bounds.lLbound);
        LONG ub = 0;
        ::SafeArrayGetUBound(pArray, 1, &ub);
        bounds.cElements = 1 + ub - bounds.lLbound;
    }
    cursor_colser   c_closer(vODBCStmt[myNo]->stmt(), true);
    auto elem = iVariant();
    try
    {
        std::vector<LONG>       errorNo;
        std::vector<VARIANT>    errorMessaged;
        SQLDiagRec<>            diagRec;
        for (ULONG i = 0; i < bounds.cElements; ++i)
        {
            auto index = static_cast<LONG>(i) + bounds.lLbound;
            ::SafeArrayGetElement(pArray, &index, &elem);
            if (elem.vt == VT_BSTR && elem.bstrVal)
            {
                auto const rc = execDirect(std::wstring(elem.bstrVal), vODBCStmt[myNo]->stmt());
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
            auto elem = iVariant(VT_I4);
            elem.lVal = c;
            return elem;
        };
        return (errorNo.size()) ?
            vec2VArray(std::vector<VARIANT>{vec2VArray(std::move(errorNo), errorNo_trans),
                vec2VArray(std::move(errorMessaged))          })
            : iVariant();
    }
    catch (const std::exception&)
    {
        return iVariant();
    }
}

// テーブル一覧
VARIANT __stdcall table_list_all(__int32 myNo, VARIANT const& schemaName) noexcept
{
    auto schema_name_b = getBSTR(schemaName);
    if (!schema_name_b || myNo < 0 || vODBCStmt_size() <= myNo)
        return iVariant();
    try
    {
        std::wstring schema_name_t{ schema_name_b };
        auto schema_name = const_cast<SQLTCHAR*>(static_cast<const SQLTCHAR*>(schema_name_t.c_str()));
        auto schema_len = static_cast<SQLSMALLINT>(schema_name_t.length());
        if (schema_len == 0)      schema_name = NULL;
        auto table_func = [=](HSTMT x) {
            return ::SQLTables(x, NULL, SQL_NTS, schema_name, schema_len, NULL, SQL_NTS, NULL, SQL_NTS);
        };
        SQLUSMALLINT columns[] = { 2, 3, 4 };
        auto scheme_name_type = catalogValue(table_func, vODBCStmt[myNo]->stmt(), columns);
        auto trans = [](std::wstring& s) { return makeVariantFromSQLType(SQL_CHAR, &s[0]); };
        std::vector<VARIANT> vec;
        vec.reserve(3);
        vec.push_back(vec2VArray(std::move(scheme_name_type[0]), trans));
        vec.push_back(vec2VArray(std::move(scheme_name_type[1]), trans));
        vec.push_back(vec2VArray(std::move(scheme_name_type[2]), trans));
        return vec2VArray(std::move(vec));
    }
    catch (const std::exception&)
    {
        return iVariant();
    }
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
VARIANT __stdcall columnAttributes_all(__int32 myNo, VARIANT const& schemaName, VARIANT const& tableName) noexcept
{
    auto schema_name_b = getBSTR(schemaName);
    auto table_Name_b = getBSTR(tableName);
    if (!schema_name_b || !table_Name_b || myNo < 0 || vODBCStmt_size() <= myNo)
        return iVariant();
    try
    {
        std::wstring schema_name_t{ schema_name_b }, table_name_t{ table_Name_b };
        auto schema_name = const_cast<SQLTCHAR*>(static_cast<const SQLTCHAR*>(schema_name_t.c_str()));
        auto table_Name = const_cast<SQLTCHAR*>(static_cast<const SQLTCHAR*>(table_name_t.c_str()));
        auto schema_len = static_cast<SQLSMALLINT>(schema_name_t.length());
        auto table_len = static_cast<SQLSMALLINT>(table_name_t.length());
        if (schema_len == 0)      schema_name = NULL;   //nullptrではない
        auto const& st = vODBCStmt[myNo]->stmt();
        struct trans {  // Workaround for VC++2013
            SQLSMALLINT x;
            trans(SQLSMALLINT i) : x(i) {   }
            VARIANT operator ()(std::wstring& s) const
            {
                return makeVariantFromSQLType(x, &s[0]);
            };
        };
        //auto trans = [](SQLSMALLINT x) {
        //    return [=](std::wstring& s)  {   return makeVariantFromSQLType(x, &s[0]);    };
        //};
        std::vector<VARIANT> vec;
        vec.reserve(7);
        {
            auto column_func = [=](HSTMT x) {
                return ::SQLColumns(x, NULL, SQL_NTS, schema_name, schema_len, table_Name, table_len, NULL, SQL_NTS);
            };
            SQLUSMALLINT columns[] = { 4, 6, 11, 9, 5, 7, 17 };
            auto column_attr = catalogValue(column_func, st, columns);
            //------------------------
            auto column_name = vec2VArray(std::move(column_attr[0]), trans(SQL_CHAR));
            auto type_name = vec2VArray(std::move(column_attr[1]), trans(SQL_CHAR));
            auto is_nullable = vec2VArray(std::move(column_attr[2]), trans(SQL_SMALLINT));
            auto Decimal_Digits = vec2VArray(std::move(column_attr[3]), trans(SQL_SMALLINT));
            auto column_size = vec2VArray(std::move(column_attr[5]), trans(SQL_INTEGER));
            auto ordinal_position = vec2VArray(std::move(column_attr[6]), trans(SQL_INTEGER));
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
            auto key_value = catalogValue(primarykeys_func, st, keycolumns);    //KEY_NAME
            auto primarykeys = vec2VArray(std::move(key_value[0]), trans(SQL_CHAR));
            vec.push_back(primarykeys);         // 6
        }
        return vec2VArray(std::move(vec));
    }
    catch (const std::exception&)
    {
        return iVariant();
    }
}

namespace 
{
    class header_getter_t    {
        std::vector<column_t::name_type>    v_colname;
    public:
        void operator()(std::vector<column_t::name_type>&   colname ,
                        std::vector<SQLSMALLINT>&   ,
                        std::vector<SQLULEN>&       ,
                        std::vector<SQLSMALLINT>&   ,
                        std::vector<SQLSMALLINT>&   ,
                        std::vector<SQLSMALLINT>&   ) noexcept
        {   v_colname = std::move(colname);     }
    };
}

// テキストファイル出力
__int32 __stdcall
textOutODBC(__int32         myNo        , 
            VARIANT const&  SQL_expr    , 
            VARIANT const&  fileName    ,
            VARIANT const&  delimiter   ,
            __int8          quote       ,
            __int8          append      ,
            __int32         Code_Page   ,
            __int32         topN        ,
            __int32         flushMB     ,
            __int8          CrLf        ) noexcept
{
    auto sql = getBSTR(SQL_expr);
    auto delim_ = getBSTR(delimiter);
    auto filename = getBSTR(fileName);
    auto codepage = static_cast<UINT>(0 <= Code_Page? Code_Page: -Code_Page);
    if ( topN < 0 )     topN = (std::numeric_limits<__int32>::max)();
    std::size_t flushByte = (flushMB <= 0)? (std::numeric_limits<std::size_t>::max)():
                                            1000000 * flushMB;
    if (!sql || !filename || !delim_ || myNo < 0 || vODBCStmt_size() <= myNo)   return 0;
    __int32 ret{0};
    try
    {
        const std::wstring delim{delim_};
        const std::wstring file_name{filename};
        std::wstring elem;
        SQLSMALLINT col_N{ 0 };
        auto init_func = [&](SQLSMALLINT c) {
            col_N = c;
        };
        auto elem_func = [&](SQLSMALLINT j, wchar_t const* str, SQLSMALLINT coltype) {
            if ( quote )            elem += L"\"";
            if (str)                elem += str;
            if ( quote )            elem += L"\"";
            if ( j < col_N - 1 )    elem += delim;
        };
        fputws_t txtOut_str{filename, Code_Page, append != 0, CrLf != 0};
        std::size_t bytes{0};
        auto add_func = [&](std::size_t x) {
            if ( topN <= static_cast<__int32>(x) )      return false;
            if ( !txtOut_str.put(elem) )
            {
                ret *= -1;
                return false;
            }
            bytes += elem.size() * sizeof(wchar_t);
            if ( flushByte < bytes )
            {
                txtOut_str.flush();
                bytes = 0;
            }
            elem.clear();
            ++ret;
            txtOut_str.lineFeed();
            return true;
        };
        auto recordLen = select_table(vODBCStmt[myNo]->stmt(),
                                      std::wstring{ sql },
                                      no_header{},
                                      init_func,
                                      elem_func,
                                      add_func);//,
        return ret;
    }
    catch (const std::exception&)
    {
        return ret;
    }
}

namespace {

    VARIANT makeVariantFromSQLType(SQLSMALLINT type, LPCOLESTR expr) noexcept
    {
        if (!expr)      return iVariant();
        switch (type)
        {
        case SQL_CHAR:      case SQL_VARCHAR:       case SQL_LONGVARCHAR:
        case SQL_WCHAR:     case SQL_WVARCHAR:      case SQL_WLONGVARCHAR:
        case SQL_BINARY:    case SQL_VARBINARY:     case SQL_LONGVARBINARY:
        {
            auto ret = iVariant(VT_BSTR);
            ret.bstrVal = ::SysAllocString(expr);
            return ret;
        }
        case SQL_SMALLINT:  case SQL_INTEGER:   case SQL_BIT:   case SQL_TINYINT:
        {
            long lOut;
            auto const vdr = ::VarI4FromStr(expr, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &lOut);
            auto ret = iVariant(VT_I4);
            ret.lVal = (vdr == S_OK)? lOut: 0;
            return ret;
        }
        case SQL_BIGINT:
        {
            LONG64  i64Out;
            auto const vdr = ::VarI8FromStr(expr, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &i64Out);
            auto ret = iVariant(VT_I8);
            ret.llVal = (vdr == S_OK)? i64Out: 0;
            return ret;
        }
        case SQL_NUMERIC:   case SQL_DECIMAL:   case SQL_FLOAT: case SQL_REAL:  case SQL_DOUBLE:
        {
            double dOut;
            auto const vdr = ::VarR8FromStr(expr, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &dOut);
            auto ret = iVariant(VT_R8);
            ret.dblVal = (vdr == S_OK)? dOut: 0.0;
            return ret;
        }
        case SQL_TYPE_DATE: case SQL_TYPE_TIME: case SQL_TYPE_TIMESTAMP:
        {
            OLECHAR date_expr[] = L"2001-01-01 00:00:00";
            auto p = expr;
            auto q = date_expr;
            while (*p != L'\0' && *p != L'.' && *q != L'\0')       *q++ = *p++;
            *q = L'\0';
            DATE dOut;
            auto const vdr = ::VarDateFromStr(date_expr, LANG_JAPANESE, LOCALE_NOUSEROVERRIDE, &dOut);
            auto ret = iVariant(VT_DATE);
            ret.date = dOut;
            return ret;
        }
        default:    return iVariant();
        }
    }

    BSTR getBSTR(VARIANT const& expr) noexcept
    {
        if (expr.vt & VT_BYREF)
            return ((expr.vt & VT_BSTR) && expr.pbstrVal) ? *expr.pbstrVal : nullptr;
        else
            return ((expr.vt & VT_BSTR) && expr.bstrVal) ? expr.bstrVal : nullptr;
    }

    BSTR getBSTR(VARIANT const* expr) noexcept
    {
        return expr? getBSTR(*expr): nullptr;
    }

    // std::vector<VARIANT> ==> Variant()
    template <typename Container_t>
    VARIANT vec2VArray(Container_t&& cont) noexcept
    {
        static_assert(!std::is_reference<Container_t>::value, "vec2VArray's parameter is a rvalue reference !!");
        auto trans = [](typename Container_t::reference x) -> typename Container_t::reference   {
            return x;
        };
        return vec2VArray(std::move(cont), trans);
    }

    // std::vector<T> ==> Variant()
    template <typename Container_t, typename F>
    VARIANT vec2VArray(Container_t&& cont, F&& trans) noexcept
    {
        static_assert(!std::is_reference<Container_t>::value, "vec2VArray's parameter is a rvalue reference !!");
        SAFEARRAYBOUND rgb = { static_cast<ULONG>(cont.size()), 0 };
        safearrayRAII pArray{::SafeArrayCreate(VT_VARIANT, 1, &rgb)};
        char* it = nullptr;
        ::SafeArrayAccessData(pArray.get(), reinterpret_cast<void**>(&it));
        if (!it)            return iVariant();
        auto const elemsize = ::SafeArrayGetElemsize(pArray.get());
        std::size_t i{0};
        try
        {
            for (auto p = cont.begin(); p != cont.end(); ++p, ++i)
                std::swap(*reinterpret_cast<VARIANT*>(it + i * elemsize), std::forward<F>(trans)(*p));
            auto ret = iVariant(VT_ARRAY | VT_VARIANT);
            ret.parray = pArray.get();
            try { cont.clear(); }
            catch (...) {}
            return ret;
        }
        catch (const std::exception&)
        {
            return iVariant();
        }
    }

    std::vector<std::vector<VARIANT>>
        selectODBC_columnWise_imple(odbc_raii_statement& st, BSTR sql, VARIANT& header)
    {
        std::vector<std::vector<VARIANT>> vec;
        auto init_func = [&](SQLSMALLINT c) {
            vec.resize(c);
        };
        auto elem_func = [&](SQLSMALLINT j, wchar_t const* str, SQLSMALLINT coltype) {
            vec[j].push_back(makeVariantFromSQLType(coltype, str));
        };
        auto add_func = [&](std::size_t) {};
        header_getter header_func;
        auto recordLen = select_table(  st,
                                    std::wstring{ sql },
                                header_func,
                            init_func,
                        elem_func,
                    add_func);
        ::VariantClear(&header);
        std::swap(header, header_func.getHeader());
        return vec;
    }

    VARIANT header_getter::getHeader() noexcept
    {
        auto bstr_trans = [](column_t::name_type& c) {
            auto elem = iVariant(VT_BSTR);
            elem.bstrVal = ::SysAllocString(c.data());
            return elem;
        };
        auto type_trans = [](SQLSMALLINT t) {
            auto elem = iVariant(VT_BSTR);
            auto const str = getTypeStr(t);
            elem.bstrVal = ::SysAllocString(str.empty() ? nullptr : &str[0]);
            return elem;
        };
        auto colname_array = vec2VArray(std::move(v_colname), bstr_trans);
        auto coltype_array = vec2VArray(std::move(v_coltype), type_trans);
        try
        {
            std::vector<VARIANT> vec{ colname_array, coltype_array };
            return vec2VArray(std::move(vec));
        }
        catch (const std::exception&)
        {
            return iVariant();
        }
    }

    //**********************************
    fputws_t::fputws_t(BSTR fileName, __int32 Code_Page, bool append, bool CrLf) noexcept
        : codepage_{Code_Page}, CrLf_{CrLf}
    {
        FILE* fp = nullptr;
        auto openmode = std::wstring(append? L"a+": L"w");
        openmode += (codepage_==1200)? L"t, ccs=UTF-16LE": L"b";
        auto err = ::_wfopen_s(&fp, fileName,  openmode.data());
        fc_RAII = fileCloseRAII{err? nullptr: fp};
        if ( codepage_ == 65001 ) //UTF-8 with BOM
        {
            char bom[] = {char(0xEF), char(0xBB), char(0xBF), '\0' };
            std::fputs(bom, fp);
        }
    }

    void fputws_t::lineFeed() const noexcept
    {
        auto fp = fc_RAII.get();
        if ( fp )
        {
            if ( codepage_ == 1200 ) {
                std::fputwc(L'\n', fp);
            } else {
                if ( CrLf_ )    std::fputc('\r', fp);
                std::fputc('\n', fp);
            }
        }
    }

    bool fputws_t::put(std::wstring const& x) const noexcept
    {
        auto fp = fc_RAII.get();
        if ( fp )
        {
            switch (codepage_)
            {
            case 1200:  // UTF-16, UTF-8
            {
                return  (0 <= std::fputws(x.data(), fp))? true: false;
            }
            case 932:  case 1252:  case 65001:  case -65001:   //SHIFT-JIS, ANSI, UTF-8, UTF-8N
            {
                auto codepage = static_cast<UINT>(0 <= codepage_? codepage_: -codepage_);
                std::string buf;
                auto b = ::WideCharToMultiByte(codepage, 0, x.data(), -1, nullptr, 0, nullptr, nullptr);
                buf.resize(b);
                b = ::WideCharToMultiByte(codepage, 0, x.data(), -1, &buf[0], b, nullptr, nullptr); 
                return  (0 <= std::fputs(buf.data(), fp))? true: false;
            }
            }
        }
        return false;
    }

    void fputws_t::flush() const noexcept
    {
        auto fp = fc_RAII.get();
        if ( fp )       std::fflush(fp);
    }

}    //namespace {

