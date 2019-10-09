//odbcResource.hpp
//Copyright (c) 2015 mmYYmmdd

#pragma once

#include <sql.h>
#include <sqlext.h>
#include <odbcinst.h>
#include <string>
#include <vector>
#include <array>
#include <type_traits>
#include <map>

#if _MSC_VER < 1900
#define noexcept throw()
#endif

#pragma comment(lib, "odbccp32.lib")

namespace mymd {

class odbc_raii_env {
    HENV    henv;
    odbc_raii_env(const odbc_raii_env&) = delete;
    odbc_raii_env(odbc_raii_env&&) = delete;
    odbc_raii_env& operator =(const odbc_raii_env&) = delete;
    odbc_raii_env& operator =(odbc_raii_env&&) = delete;
public:
    odbc_raii_env() noexcept;
    ~odbc_raii_env() noexcept;
    bool AllocHandle() noexcept;
    template <typename T>
    RETCODE invoke(T&& expr) const noexcept
    {
        return (std::forward<T>(expr))(henv);
    }
};

//**************************************************************
class odbc_raii_connect {
    HDBC    hdbc;
    bool    autoCommit;
    odbc_raii_connect(const odbc_raii_connect&) = delete;
    odbc_raii_connect(odbc_raii_connect&&) = delete;
    odbc_raii_connect& operator =(const odbc_raii_connect&) = delete;
    odbc_raii_connect& operator =(odbc_raii_connect&&) = delete;
public:
    odbc_raii_connect() noexcept;
    ~odbc_raii_connect() noexcept;
    bool AllocHandle(const odbc_raii_env& env) noexcept;
    void set_autoCommit(bool) noexcept;
    bool rollback() const noexcept;
    bool commit() const noexcept;
    template <typename T>
    RETCODE invoke(T&& expr) const noexcept
    {
        return (std::forward<T>(expr))(hdbc);
    }
};

//**************************************************************
class odbc_raii_statement {
    HSTMT   hstmt;
    odbc_raii_statement(const odbc_raii_statement&) = delete;
    odbc_raii_statement(odbc_raii_statement&&) = delete;
    odbc_raii_statement& operator =(const odbc_raii_statement&) = delete;
    odbc_raii_statement& operator =(odbc_raii_statement&&) = delete;
public:
    odbc_raii_statement() noexcept;
    ~odbc_raii_statement() noexcept;
    std::wstring AllocHandle(const std::wstring& connectName, const odbc_raii_connect& con);
    template <typename T>
    RETCODE invoke(T&& expr) const noexcept
    {
        return (std::forward<T>(expr))(hstmt);
    }
};

//********************************************
// RAII for Cursor      (SQLCloseCursor)
//********************************************
class cursor_colser {
    const odbc_raii_statement&  h_;
    bool close_;
public:
    cursor_colser(const odbc_raii_statement& h, bool b) noexcept;
    ~cursor_colser() noexcept;
};

//************************************************************************************
// DB connection packgage (odbc_raii_env + odbc_raii_connect + odbc_raii_statement)
// まずこれを作ってDBに接続する
//************************************************************************************
class odbc_set {
    odbc_raii_env       env;
    odbc_raii_connect   con;
    odbc_raii_statement st;
    std::wstring        errorMessage_;
public:
    explicit odbc_set(const std::wstring& connectName,
                      decltype(SQL_CURSOR_FORWARD_ONLY) cursor_type = SQL_CURSOR_FORWARD_ONLY) noexcept;
    odbc_raii_connect& conn() noexcept;
    odbc_raii_statement& stmt() noexcept;
    void set_autoCommit(bool) noexcept;
    bool rollback() const noexcept;
    bool commit() const noexcept;
    void set_cursor_type(decltype(SQL_CURSOR_STATIC) cursor_type) const noexcept;
    bool isError() const noexcept;
    std::wstring errorMessage() const;
};

//**************************************************************
std::wstring getTypeStr(SQLSMALLINT) noexcept;
//********************************************************

struct column_t {
    static std::size_t const nameSize = 256;
    using name_type = std::array<wchar_t, nameSize>;
    static std::size_t const bufferSize = 65536;
};
//********************************************************

//Diagnostic Message 診断メッセージ
template <SQLSMALLINT HandleType = SQL_HANDLE_STMT, std::size_t bufferSize = SQL_MAX_MESSAGE_LENGTH>
class SQLDiagRec {
    SQLSMALLINT recNum;
    SQLWCHAR SQLState[6];
    SQLWCHAR szErrorMsg[bufferSize];
public:
    SQLDiagRec() noexcept : recNum{ 1 }
    {
        SQLState[0] = L'\0'; szErrorMsg[0] = L'\0';
    }
    void setnum(SQLSMALLINT a) noexcept { recNum = a; }
    std::wstring getMessage() const          { return szErrorMsg; }
    std::wstring getState() const            { return SQLState; }
    RETCODE operator ()(HSTMT x) noexcept
    {
        SQLSMALLINT o_o;
        return ::SQLGetDiagRec(HandleType, x, recNum, SQLState, NULL, szErrorMsg, bufferSize, &o_o);
    }
};

//********************************************************
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
// カタログ関数  Catalog Infomation
template <typename FC, typename Iter_t>
std::vector<std::vector<std::wstring>>
catalogValue(FC&&                           catalog_func    ,
             odbc_raii_statement const&     st              ,
             Iter_t                         columnNumber_begin  ,
             Iter_t                         columnNumber_end    )
{
    std::vector<std::vector<std::wstring>> ret(columnNumber_end - columnNumber_begin);
    auto result = st.invoke(std::forward<FC>(catalog_func));
    if (SQL_SUCCESS != result)      return ret;
    SQLSMALLINT nresultcols{ 0 };
    {
        auto pl = &nresultcols;
        auto result = st.invoke(
            [=](HSTMT x) { return ::SQLNumResultCols(x, pl); }
        );
    }
    const std::size_t ColumnNameLen = 1024;
    std::vector<std::array<wchar_t, ColumnNameLen>> buffer_v(nresultcols);
    SQLLEN   pcbValue{ 0 };
    auto p_pcbValue = &pcbValue;
    cursor_colser   c_closer{st, true};
    auto buffer_v_it = buffer_v.begin();
    for (auto j = 0; j < nresultcols; ++j, ++buffer_v_it)
    {
        auto p_rgbValue = static_cast<SQLPOINTER>(buffer_v_it->data());
        auto result = st.invoke(
            [=](HSTMT x) { return ::SQLBindCol(x,
                                            j+1,
                                        SQL_C_WCHAR,
                                    p_rgbValue,
                                ColumnNameLen,
                            p_pcbValue);
        }
        );
        if (SQL_SUCCESS != result)      return ret;
    }
    auto SQLFetch_expr = [=](HSTMT x) { return ::SQLFetch(x); };
    while (true)
    {
        auto fetch_result = st.invoke(SQLFetch_expr);
        if ((SQL_SUCCESS != fetch_result) && (SQL_SUCCESS_WITH_INFO != fetch_result))
            break;
        for (auto p = columnNumber_begin; p < columnNumber_end; ++p)
        {
            auto num = (*p) - 1;
            if (num < 0 || nresultcols <= num)    continue;
            ret[p-columnNumber_begin].push_back(buffer_v[num].data());
            buffer_v[num].fill(wchar_t{});
        }
    }
    return ret;
}

template <typename FC, typename Arr>
std::vector<std::vector<std::wstring>>
catalogValue(FC&&                       catalog_func,
             odbc_raii_statement const& st          ,
             Arr&&                      arr         )
{
    return catalogValue(std::forward<FC>(catalog_func)  ,
                        st  ,
                        std::begin(std::forward<Arr>(arr))  ,
                        std::end(std::forward<Arr>(arr))    );
}

//******************************************************************
// SQLExecDirect
//  SELECT  ,  INSERT  ,  UPDATE  ,  ...
RETCODE execDirect(const std::wstring& sql_expr, const odbc_raii_statement& stmt) noexcept;

//******************************************************************

    // 以下は bindParameters_exec のための
namespace detail    {

    // https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlparamdata-function?redirectedfrom=MSDN&view=sql-server-2017
    // https://docs.microsoft.com/ja-jp/sql/odbc/reference/develop-app/using-arrays-of-parameters?view=sql-server-2017

    class ValuePtrPtr  {
        SQLPOINTER          current;   //char*, wchar_t*
        std::vector<int>    actlen;
        std::size_t         counter;
    public:
        ValuePtrPtr() : current{nullptr}, counter{0}     {}
        void reserve(std::size_t s) &   { actlen.reserve(s); }
        void push_back(int len) &       { actlen.push_back(len); }
        void bind_late(SQLPOINTER p) &  { current = p; }
        SQLPOINTER get_current() const  { return 0 < actlen[counter]? current: nullptr; }
        SQLLEN get_size() const         { return 0 < actlen[counter]? static_cast<SQLLEN>(actlen[counter]): 0; }
        ValuePtrPtr& operator++() &
        {
            current = static_cast<char*>(current) + (0 < actlen[counter]? actlen[counter]: 0);
            ++counter;
            return *this;
        }
    };

    struct value_container_base    {
        virtual ~value_container_base() {};
        virtual SQLLEN buff_len() const {   return 0;   }
        virtual void move_vpp(std::map<SQLPOINTER, ValuePtrPtr>&) { }
    };

    template <typename value_type, typename = void>
    struct value_container : value_container_base   {
        virtual ~value_container() = default;
        std::vector<value_type> hold;
        std::vector<SQLLEN>     StrLen_or_IndPtr;
        void reserve(std::size_t s)     { hold.reserve(s); StrLen_or_IndPtr.reserve(s); }
        template <typename pointer_type>
        void push_back(pointer_type p)
        {
            hold.push_back(p? *p: value_type{});
            StrLen_or_IndPtr.push_back(p? sizeof(value_type): SQL_NULL_DATA);
        }
        SQLPOINTER  begin1()        {   return &hold[0];   }
        SQLLEN*     begin2()        {   return &StrLen_or_IndPtr[0];    }
    };

    template <typename C, std::size_t N>
    struct value_container<std::array<C, N>> : value_container_base   {
        virtual ~value_container() = default;
        std::vector<std::array<C, N>>   hold;
        std::vector<SQLLEN>             StrLen_or_IndPtr;
        void reserve(std::size_t s)     { hold.reserve(s); StrLen_or_IndPtr.reserve(s); }
        template <typename pointer_type>
        void push_back(pointer_type p)
        {
            hold.push_back(p? *p: std::array<C, N>{});
            StrLen_or_IndPtr.push_back(p? SQL_NTS: SQL_NULL_DATA);
        }
        SQLPOINTER  begin1()        {   return &hold[0];   }
        SQLLEN*     begin2()        {   return &StrLen_or_IndPtr[0];    }
        SQLLEN buff_len() const override    {   return sizeof(std::array<C, N>);    }
    };

    template <typename STR>
    struct value_container<STR, std::void_t<typename STR::traits_type>> : value_container_base  {
        virtual ~value_container() = default;
        STR                 hold;
        std::vector<SQLLEN> StrLen_or_IndPtr;
        ValuePtrPtr         vpp;
        void reserve(std::size_t s)     { hold.reserve(s); StrLen_or_IndPtr.reserve(s); vpp.reserve(s); }
        template <typename pointer_type>
        void push_back(pointer_type p)
        {
            if (p) hold.append(*p);
            auto len = p? static_cast<SQLLEN>(sizeof(STR::value_type) * p->size()): 0;
            StrLen_or_IndPtr.push_back(SQL_LEN_DATA_AT_EXEC(len));
            vpp.push_back(p? static_cast<int>(len): -1);
        }
        SQLPOINTER  begin1()        { return &hold[0]; }
        SQLLEN*     begin2()        { return &StrLen_or_IndPtr[0]; }
        void move_vpp(std::map<SQLPOINTER, ValuePtrPtr>& m) override
        {
            vpp.bind_late(&hold[0]);
            m.insert_or_assign(&hold[0], std::move(vpp));
        }
    };

    //  numeric(13, 0) なら {SQL_C_CHAR, SQL_NUMERIC, 13, 0}
    struct bindParameterAttribute {
        SQLSMALLINT       ValueType;           // fCType
        SQLSMALLINT       ParameterType;       // fSqlType 
        SQLULEN           ColumnSize;          // cbColDef 
        SQLSMALLINT       DecimalDigits;
    };

    template <typename...Container_t>
    RETCODE bindParameters_imple(HSTMT,
                                 std::size_t,
                                 SQLUSMALLINT,
                                 bindParameterAttribute*,
                                 std::vector<std::shared_ptr<value_container_base>>&,
                                 Container_t&&...) noexcept
    {
        return SQL_SUCCESS;
    }

    template <typename Container_0, typename... Container_t>
    RETCODE bindParameters_imple(HSTMT                              h,
                                 std::size_t                        size,                             
                                 SQLUSMALLINT                       ParameterNumber,
                                 bindParameterAttribute*            attr,
                                 std::vector<std::shared_ptr<value_container_base>>& value_container_vec,
                                 Container_0&&                      container0,
                                 Container_t&&...                   containers  ) noexcept
    {
        if (size !=  std::size(std::forward<Container_0>(container0)))
            return SQL_ERROR;
        using value_type = std::remove_cv_t<std::remove_reference_t<decltype(*container0[0])>>;
        auto value_hdr = std::make_shared<value_container<value_type>>();
        value_hdr->reserve(size);
        for (auto const& p_elem : std::forward<Container_0>(container0) )
        {
            value_hdr->push_back(p_elem);
        }
        auto rt = ::SQLBindParameter(h,
                                     ParameterNumber,
                                     SQL_PARAM_INPUT,
                                     attr->ValueType,
                                     attr->ParameterType,
                                     attr->ColumnSize, 
                                     attr->DecimalDigits,
                                     value_hdr->begin1(),
                                     value_hdr->buff_len(),     //BufferLength
                                     value_hdr->begin2()  );
        value_container_vec.push_back(std::move(value_hdr));
#ifndef NDEBUG
        if  ( rt != SQL_SUCCESS )
        {
            mymd::SQLDiagRec<> diagRec;
            diagRec(h);
            auto ms = diagRec.getMessage();
            return rt;
        }
#endif
        return (rt != SQL_SUCCESS)? rt: bindParameters_imple(h,
                                                             size,
                                                             ParameterNumber + 1,
                                                             attr + 1,
                                                             value_container_vec,
                                                             std::forward<Container_t>(containers)...);
    }

}       //</detail>


//  https://docs.microsoft.com/ja-jp/sql/odbc/reference/appendixes/sql-data-types?view=sql-server-2017
//  https://docs.microsoft.com/ja-jp/sql/odbc/reference/appendixes/c-data-types?view=sql-server-2017
//  https://docs.microsoft.com/ja-jp/sql/t-sql/data-types/decimal-and-numeric-transact-sql?view=sql-server-2017

// *********************************************************************************
// パラメーター配列を使用した実行
// https://docs.microsoft.com/ja-jp/sql/odbc/reference/syntax/sqlbindparameter-function?view=sql-server-2017
// Container_0, Container_t... は std::optional（またはポインタ）のシーケンスを想定
// *********************************************************************************

//bindParameterAttribute
struct colPosition_parameterCType   {
    int             colPosition;        // column number(1 origin)
    SQLSMALLINT     parameterCType;     // C data type
};

template <typename Container_0, typename... Container_t>
RETCODE bindParameters_exec(const odbc_raii_statement&          stmt        ,
                            std::wstring                        schema_name ,
                            std::wstring                        table_name  ,
                            std::wstring const&                 execSQL_expr,
                            colPosition_parameterCType* const   col_ctype   ,       // bindParameterAttribute*
                            Container_0&&                       container0  ,
                            Container_t&&...                    containers  ) noexcept
{
    try
    {
        auto const size = std::size(std::forward<Container_0>(container0));
        if (size == 0 || schema_name.empty() || table_name.empty())     return SQL_ERROR; 
        cursor_colser   c_closer{stmt, true};
        // 列属性を取得 get columns attribute
        detail::bindParameterAttribute  attr[1+sizeof...(containers)];
        {
            auto column_func = [&](HSTMT x) {
                auto sdfsd = &schema_name[0];
                return ::SQLColumns(x,
                                    NULL,
                                    SQL_NTS,
                                    &schema_name[0],
                                    static_cast<SQLSMALLINT>(schema_name.length()),
                                    &table_name[0],
                                    static_cast<SQLSMALLINT>(table_name.length()),
                                    NULL,
                                    SQL_NTS);
            };
            SQLUSMALLINT columns4[] = {17, 5, 7, 9};    //ORDINAL_POSITION, DATA_TYPE, COLUMN_SIZE, DECIMAL_DIGITS
            auto column_attr = catalogValue(column_func, stmt, columns4);
            std::vector<int>  colPosition(column_attr[0].size());
            std::transform(column_attr[0].begin(), column_attr[0].end(), colPosition.begin(), [](std::wstring const& w){ return std::stoi(w); });
            for ( auto i = 0 ; i < 1+sizeof...(containers); ++i )
            {
                auto ppos = std::find(colPosition.begin(), colPosition.end(), col_ctype[i].colPosition);
                if (ppos < colPosition.end())
                {
                    auto pos = ppos - colPosition.begin();
                    attr[i].ValueType = col_ctype[i].parameterCType;
                    if ( column_attr[1][pos].empty() || column_attr[2][pos].empty() )   return SQL_ERROR;
                    attr[i].ParameterType = static_cast<SQLSMALLINT>(std::stoi(column_attr[1][pos]));
                    attr[i].ColumnSize = static_cast<SQLULEN>(std::stoi(column_attr[2][pos]));
                    attr[i].DecimalDigits = static_cast<SQLSMALLINT>(std::stoi(column_attr[3][pos].empty()? L"0":column_attr[3][pos]));
                }
                else    return SQL_ERROR;
            }
        }   // </列属性を取得 get columns attribute>
        HSTMT h;
        stmt.invoke([&](HSTMT x){ h = x;  return SQL_SUCCESS;});
        ::SQLSetStmtAttr(h, SQL_ATTR_PARAM_BIND_TYPE, SQL_PARAM_BIND_BY_COLUMN, 0);
        ::SQLSetStmtAttr(h, SQL_ATTR_PARAMSET_SIZE, reinterpret_cast<SQLPOINTER>(size), 0);
        SQLRETURN  rt;
        std::vector<std::shared_ptr<detail::value_container_base>> value_container_vec;
        value_container_vec.reserve(1+ sizeof...(containers));
        rt = detail::bindParameters_imple(h,
                                          size,
                                          1,
                                          attr,
                                          value_container_vec,
                                          std::forward<Container_0>(container0),
                                          std::forward<Container_t>(containers)...);
        if (SQL_SUCCESS != rt && SQL_SUCCESS_WITH_INFO != rt)       return rt;
        std::map<SQLPOINTER, detail::ValuePtrPtr>   vpp_map;
        for ( auto& elem : value_container_vec )    elem->move_vpp(vpp_map);
        if ( SQL_NEED_DATA == (rt = execDirect(execSQL_expr, stmt)) )
        {
            SQLPOINTER   ValuePtr;
            while ( SQL_NEED_DATA ==::SQLParamData(h, &ValuePtr) )
            {
                if ( vpp_map.count(ValuePtr) )
                {
                    detail::ValuePtrPtr& vpp = vpp_map[ValuePtr];
                    auto rt2 = ::SQLPutData(h, vpp.get_current(), vpp.get_size());
                    if (SQL_SUCCESS != rt2 && SQL_SUCCESS_WITH_INFO != rt2)   return rt2;
                    ++vpp;
                }
            }
        }
        return rt;
    }
    catch(...)
    {
        return SQL_ERROR;
    }
}

//******************************************************************
// 列の属性取得
//******************************************************************
template <typename F>
SQLSMALLINT columnAttribute(odbc_raii_statement const&          stmt,
                            std::wstring const&                 sql_expr,
                            std::vector<std::wstring>*          pBuffer,
                            std::vector<SQLLEN>*                pdatastrlen,
                            F&&                                 write_func,
                            bool                                cursor_close    ) noexcept
{
    auto const rc = execDirect(sql_expr, stmt);
    if (rc == SQL_ERROR || rc == SQL_INVALID_HANDLE)    return 0;
    SQLSMALLINT nresultcols{ 0 };
    {
        auto pl = &nresultcols;
        auto const rc = stmt.invoke(
            [=](HSTMT x) { return ::SQLNumResultCols(x, pl); }
        );
        if (SQL_SUCCESS != rc)  return 0;
    }
    try
    {
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
        auto SQLDescribeColExpr = [&](HSTMT x) {
            return ::SQLDescribeCol(x,
                                static_cast<UWORD>(j+1),
                            colname[j].data(),
                        static_cast<SQLSMALLINT>(column_t::nameSize * sizeof(wchar_t)),
                        &colnamelen[j],
                        &coltype[j],
                            &collen[j],
                                &scale[j],
                                    &nullable[j]);
        };
        auto SQLBindColExpr = [&](HSTMT x) {
            return ::SQLBindCol(x,
                                static_cast<UWORD>(j+1),
                                SQL_C_WCHAR,
                                &(*pBuffer)[j][0],
                                (*pBuffer)[j].size() * sizeof(wchar_t),
                                &(*pdatastrlen)[j]);
        };
        auto const StrSizeofColumn = column_t::bufferSize;
        cursor_colser   c_closer{ stmt, cursor_close };
        for (j = 0; j < nresultcols; ++j)
        {
            auto rc = stmt.invoke(SQLDescribeColExpr);
            if (pBuffer && pdatastrlen)
            {
                auto dlen = collen[j];
                (*pBuffer)[j].resize((0 < dlen && dlen+4 < StrSizeofColumn) ? dlen+4 : StrSizeofColumn);
                rc = stmt.invoke(SQLBindColExpr);
            }
        }
        std::forward<F>(write_func)(colname, colnamelen, collen, nullable, coltype, scale);
        return nresultcols;
    }
    catch (const std::exception&)
    {
        return 0;
    }
}

//************************************************************************************
//  ヘッダ情報が不要の場合のファンクタ  Functor when header information is not required
//************************************************************************************
struct no_header {
    void operator()(std::vector<column_t::name_type>&,
                    std::vector<SQLSMALLINT>&,
                    std::vector<SQLULEN>&,
                    std::vector<SQLSMALLINT>&,
                    std::vector<SQLSMALLINT>&,
                    std::vector<SQLSMALLINT>&) const noexcept   {   }
    void operator()(SQLSMALLINT) const noexcept                 {   }   //何もしないinit_func
};

struct bool_sentinel {
    explicit operator bool() const noexcept { return true; }
    friend bool operator ,(bool b, const bool_sentinel&) noexcept { return b; }
};

//******************************************************************
template<std::size_t N> struct uint_type_ { using type = std::size_t; };
template<> struct uint_type_<1> { using type = std::uint8_t; };
template<> struct uint_type_<2> { using type = std::uint16_t; };
template<> struct uint_type_<4> { using type = std::uint32_t; };
template<> struct uint_type_<8> { using type = std::uint64_t; };

template <typename F, typename U = void>
struct counter_type_for_ {
    using type = std::size_t;
};

template <typename R, typename T>
struct counter_type_for_<R(*)(T), void> {
    using type = typename uint_type_<sizeof(T)>::type;
};

template <typename R, typename T>
struct counter_type_for_<R(&)(T), void> {
    using type = typename uint_type_<sizeof(T)>::type;
};

template <typename R, typename C, typename T> auto memf_param_type(R(C::*)(T))->T;

template <typename R, typename C, typename T> auto memf_param_type(R(C::*)(T) const)->T;

template <typename F>
struct counter_type_for_ <F, std::void_t<decltype(memf_param_type(&F::operator()))>> {
    using type = typename uint_type_<sizeof(decltype(memf_param_type(&F::operator())))>::type;
};

template <typename F>
using counter_type_for = typename counter_type_for_<F>::type;

//********************************************************************************
//  Query(SELECT)    SELECT 文の実行
//********************************************************************************
/* void header_func(std::vector<column_t::name_type>&   column_name,
                    std::vector<SQLSMALLINT>&           column_name_len,
                    std::vector<SQLULEN>&               column_len,
                    std::vector<SQLSMALLINT>&           nullable,
                    std::vector<SQLSMALLINT>&           column_type,
                    std::vector<SQLSMALLINT>&           scale       );

   void init_func(SQLSMALLINT number_of_columns);

   void elem_func(SQLSMALLINT column_index, wchar_t const* data_str, SQLSMALLINT coltype);

   void/bool add_func(counter_t record_index);   counter_t is any integer type
   
   return  :  number of records    (type of parameter of add_func)     */
//********************************************************************************
template <typename FH, typename FI, typename FE, typename FA, typename FtCur>
auto select_table(odbc_raii_statement const& stmt           ,
                  std::wstring const&         sql_expr      ,
                  FH&&                        header_func   ,
                  FI&&                        init_func     ,
                  FE&&                        elem_func     ,
                  FA&&                        add_func      ,
                  FtCur&&                     fetch_expr    ) noexcept -> counter_type_for<FA>
{
    try
    {
        std::vector<SQLSMALLINT>        coltype;
        auto write_func = [&](std::vector<column_t::name_type>&   colname_,
                            std::vector<SQLSMALLINT>&           colnamelen_,
                        std::vector<SQLULEN>&               collen_,
                    std::vector<SQLSMALLINT>&           nullable_,
                std::vector<SQLSMALLINT>&           coltype_,
            std::vector<SQLSMALLINT>&           scale_)
        {
            coltype = coltype_;
            std::forward<FH>(header_func)(colname_, colnamelen_, collen_, nullable_, coltype_, scale_);
        };
        std::vector<std::wstring>           buffer;
        std::vector<SQLLEN>                 datastrlen;
        auto nresultcols = columnAttribute( stmt        ,
                                            sql_expr    ,
                                            &buffer     ,
                                            &datastrlen ,
                                            write_func  ,
                                            false       );
        if (nresultcols == 0)          return 0;
        //-----------------------------------------------
        cursor_colser   c_closer{ stmt, true };
        bool_sentinel   bp;
        if (!(std::forward<FI>(init_func)(nresultcols), bp))    return 0;
        counter_type_for<FA> counter{ 0 };
        while (true)
        {
            for (int j = 0; j < nresultcols; ++j)
                buffer[j][0] = L'\0';
            auto const rc = stmt.invoke(std::forward<FtCur>(fetch_expr));
            if (rc == SQL_SUCCESS || rc == SQL_SUCCESS_WITH_INFO)
            {
                for (SQLSMALLINT j = 0; j < nresultcols; ++j)
                {
                    std::forward<FE>(elem_func)(j,
                                    (SQL_NULL_DATA == datastrlen[j]) ? nullptr : buffer[j].data(),
                                coltype[j]);
                }
                if (!(std::forward<FA>(add_func)(counter++), bp))
                    break;
            }
            else
            {
                break;
            }
        }
        return counter;
    }
    catch (const std::exception&)
    {
        return 0;
    }
}

//*****************************************************
// usual case of SELECT      (FtCur : SQLFetch)
//*****************************************************
template <typename FH, typename FI, typename FE, typename FA>
auto select_table(odbc_raii_statement const& stmt           ,
                  std::wstring const&        sql_expr       ,
                  FH&&                        header_func   ,
                  FI&&                        init_func     ,
                  FE&&                        elem_func     ,
                  FA&&                        add_func      ) noexcept-> counter_type_for<FA>
{
    return select_table(stmt    ,
                        sql_expr    ,
                        std::forward<FH>(header_func)   ,
                        std::forward<FI>(init_func)     ,
                        std::forward<FE>(elem_func)     ,
                        std::forward<FA>(add_func)      ,
                        [](HSTMT x) { return ::SQLFetch(x); }   );
}

}   // namespace mymd
