//odbcSample.cpp
//Copyright (c) 2017 mmYYmmdd
#include "stdafx.h"
#include "odbcResource.hpp"

using namespace mymd;

int main()
{
    odbc_set o_o{tstring{
        _T("Driver={SQL Server Native Client 11.0}; Trusted_Connection=YES; Server=MY-PC\\SQLEXPRESS; DATABASE=sampleDB;")
    }};
    std::vector<std::vector<tstring>> vec;
    std::vector<tstring> elem;
    SQLSMALLINT col_N{ 0 };
    auto init_func = [&](SQLSMALLINT c) {
        elem.resize(col_N = c);
    };
    auto elem_func = [&](SQLSMALLINT j, TCHAR const* str, SQLSMALLINT coltype) {
        if (str)    elem[j] = str;
        else        elem[j] = _T("NULL");
    };
    auto add_func = [&](std::size_t x) {
        vec.push_back(std::move(elem));
        elem.resize(col_N);
    };
    auto recordLen = select_table(o_o.stmt(),
        tstring{ _T("SELECT * FROM myTable") },
        no_header{},
        init_func,
        elem_func,
        add_func);
    return 0;
}
