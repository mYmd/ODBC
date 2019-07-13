//csvmap.hpp
//Copyright (c) 2019 mmYYmmdd

//WINDOWS ONLY
#pragma once
#include <sstream>
#include <stringapiset.h>

//CSVをマップする
//ユーザーが使う関数は mymd::map_csv

namespace mymd  {

inline
void code_convert_append(std::wstring& target, char const* begin, char const* end, UINT codepage)
{
    if (begin < end)
    {
        auto const targetlen = target.length();
        target.resize(targetlen + end - begin, L'\0');
        auto re = ::MultiByteToWideChar(codepage,
                                        MB_ERR_INVALID_CHARS,
                                        begin,
                                        static_cast<int>(end - begin),
                                        &target[targetlen],
                                        static_cast<int>(1 + end - begin));
        target.resize(targetlen + (re? re: 0));
    }
}

inline
void code_convert_append(std::wstring& target, std::string const& source, UINT codepage)
{
    code_convert_append(target, source.data(), source.data() + source.length(), codepage);
}

inline
void code_convert_append(std::wstring& target, std::wstring const& source, UINT)
{
    target += source;
}

inline
void code_convert_append(std::string& target, std::string const& source, UINT)
{
    target += source;
}

namespace detail {

    inline char    quote_(char)         { return '"'; }
    inline wchar_t quote_(wchar_t)      { return L'"'; }

    inline char    endl_(char)          { return '\n'; }
    inline wchar_t endl_(wchar_t)       { return L'\n'; }

    template<typename R>
    std::size_t quote_count(std::basic_string<R> const & buf)
    {
        const R quote{quote_(R{})};
        std::size_t q_count{0};
        for (auto i = buf.cbegin(); i < buf.cend(); ++i)
            if (*i == quote)    ++q_count;
        return q_count;
    }

    template <typename R>
    std::basic_string<R> make_elem(typename std::basic_string<R>::const_iterator b,
                                   typename std::basic_string<R>::const_iterator e,
                                   R    quote)
    {
        std::basic_string<R>    w(b, e);
        auto wlen = w.length();
        if (2 <= wlen)
        {
            if (w[0] == quote)
            {
                w.erase(wlen-1, 1);
                w.erase(0, 1);
            }
            auto i = static_cast<int>(wlen - 2);
            while (0 <= i)
            {
                if (w[i] == quote)
                {
                    if (w[i+1] == quote)
                    {
                        w.erase(i+1, 1);
                        wlen -= 1;
                        i -= 2;
                    }
                    else    i -= 1;
                }
                else    i -= 2;
            }
        }
        return w;
    };

    //-------------------------------------------
    template<typename R, typename F>
    std::size_t map_csv_imple_elem(std::basic_string<R> const& buf, R delimiter, F&& func)
    {
        const R quote{detail::quote_(R{})};
        std::size_t count{0}, q_count{0};
        auto i = buf.cbegin();
        auto j = i;
        while ( j != buf.cend() )
        {
            if (*j == quote)
            {
                ++q_count;
                ++j;
            }
            else if (*j == delimiter && 0 == (q_count % 2))
            {
                //コールバック関数に1要素を渡す
                std::forward<F>(func)(count, make_elem(i, j, quote));
                i = j + 1;
                j = i;
                ++count;
            }
            else
                ++j;
        }
        //コールバック関数に1要素を渡す
        std::forward<F>(func)(count, make_elem(i, j, quote));
        return count + 1;
    }

    template<typename R, typename S, typename Traits, typename F>
    std::size_t map_csv_imple(std::basic_istream<S, Traits>&    stream,
                              R                     delimiter,
                              F&&                   func,
                              std::basic_string<R>& buf,
                              std::basic_string<R>& buf2,
                              std::basic_string<S>& tmp,
                              UINT                  codepage)
    {
        std::getline(stream, tmp);
        buf.clear();
        code_convert_append(buf, tmp, codepage);
        std::size_t q_count = quote_count(buf);
        while ((q_count % 2) && stream.good())
        {
            std::getline(stream, tmp);
            buf2.clear();
            code_convert_append(buf2, tmp, codepage);
            buf += endl_(R{});
            buf += buf2;
            q_count += quote_count(buf2);
        }
        return map_csv_imple_elem(buf, delimiter, std::forward<F>(func));
    }

}

//ユーザーが使う関数
// stream : 入力ストリーム（fstream または stringstream を想定 ）
// R : 出力する文字型（wchar_tを想定）
// S : 対象streamの文字型（ファイルをcodeconvertする前提ならstringを想定）
// delimiter : 区切り文字（これによって型R が決定される）
// elem_func   : 要素毎のコールバック　　[&](std::size_t count, std::wstring&& expr) 等、 (項目番号, 要素文字列)
// record_func : 1レコード読終時のコールバック　[&](std::size_t count, std::size_t size)　等、(行番号, 当該行の項目数)
// codepage : CP_UTF8, CP_ACP など
// 返り値 : 読んだ行数（テキストファイルの場合は最後の空行もカウント）

template<typename R, typename S, typename Traits, typename EF, typename RF>
std::size_t map_csv(std::basic_istream<S, Traits>& stream, R delimiter, EF&& elem_func, RF&& record_func, UINT codepage)
{
    std::basic_string<R> buf, buf2;
    std::basic_string<S> tmp;
    std::size_t rcount{0};
    while (stream.good())
    {
        auto size = detail::map_csv_imple(stream, delimiter, std::forward<EF>(elem_func), buf, buf2, tmp, codepage);
        //コールバック（行読終）
        auto b = std::forward<RF>(record_func)(rcount++, size);
        if (!b) break;
    }
    return rcount;
}

}
