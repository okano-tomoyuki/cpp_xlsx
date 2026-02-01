#include <filesystem>
#include <windows.h>
#include <ole2.h>
#include <stdio.h>
#include <stdlib.h>
#include <stdexcept>
#include "cpp_xlsx.hpp"

namespace cpp_xlsx
{

    class VariantConverter
    {
    public:
        static Value fromVariant(const VARIANT &v);
        static VARIANT toVariant(const Value &v);
    };

}

using namespace cpp_xlsx;

namespace
{

    class SingleCall
    {
    public:
        static IDispatch *call()
        {
            static SingleCall instance;
            return instance.app_;
        }

        ~SingleCall()
        {
            if (app_)
            {
                app_->Release();
                app_ = nullptr;
            }
            CoUninitialize();
        }

    private:
        IDispatch *app_;

        SingleCall()
            : app_(nullptr)
        {
            CoInitialize(NULL);

            CLSID clsid;
            HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
            if (FAILED(hr))
            {
                throw std::runtime_error("CLSIDFromProgID() failed.");
            }

            hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&app_);
            if (FAILED(hr))
            {
                throw std::runtime_error("Excel not registered properly.");
            }
        }
    };

}

Value::Value(int v) : data_(v) {}
Value::Value(double v) : data_(v) {}
Value::Value(bool v) : data_(v) {}
Value::Value(const wchar_t *s) : data_(std::wstring(s)) {}
Value::Value(const std::wstring &s) : data_(s) {}
Value::Value(const DispatchWrapper &d) : data_(d) {}
Value::Value(const Array &arr) : data_(arr) {}

std::wstring Value::toString() const
{
    // 空
    if (std::holds_alternative<std::monostate>(data_))
    {
        return L"";
    }

    // int
    if (std::holds_alternative<int>(data_))
    {
        return std::to_wstring(std::get<int>(data_));
    }

    // double
    if (std::holds_alternative<double>(data_))
    {
        return std::to_wstring(std::get<double>(data_));
    }

    // bool
    if (std::holds_alternative<bool>(data_))
    {
        return std::get<bool>(data_) ? L"true" : L"false";
    }

    // string
    if (std::holds_alternative<std::wstring>(data_))
    {
        return std::get<std::wstring>(data_);
    }

    // DispatchWrapper（デバッグ用）
    if (std::holds_alternative<DispatchWrapper>(data_))
    {
        const auto &d = std::get<DispatchWrapper>(data_);
        wchar_t buf[64];
        swprintf(buf, 64, L"[DispatchWrapper %p]", d.raw());
        return buf;
    }

    // 2D Array
    if (std::holds_alternative<Array>(data_))
    {
        const auto &arr = std::get<Array>(data_);

        std::wstring out;
        for (size_t i = 0; i < arr.size(); i++)
        {
            for (size_t j = 0; j < arr[i].size(); j++)
            {
                out += arr[i][j].toString();
                if (j + 1 < arr[i].size())
                    out += L", ";
            }
            if (i + 1 < arr.size())
                out += L"\n";
        }
        return out;
    }

    return L"";
}

// --- VARIANT → Value ---
Value VariantConverter::fromVariant(const VARIANT &v)
{
    if (v.vt == VT_EMPTY || v.vt == VT_NULL)
    {
        return Value();
    }
    else if (v.vt == VT_I4)
    {
        return Value(static_cast<int>(v.lVal));
    }
    else if (v.vt == VT_R8)
    {
        return Value(v.dblVal);
    }
    else if (v.vt == VT_BOOL)
    {
        return Value(v.boolVal == VARIANT_TRUE);
    }
    else if (v.vt == VT_BSTR)
    {
        return Value(std::wstring(v.bstrVal ? v.bstrVal : L""));
    }
    else if (v.vt == VT_DISPATCH)
    {
        return Value(DispatchWrapper(v.pdispVal));
    }
    else if (v.vt == (VT_ARRAY | VT_VARIANT))
    {
        SAFEARRAY *sa = v.parray;

        LONG lBound1, uBound1;
        LONG lBound2, uBound2;

        SafeArrayGetLBound(sa, 1, &lBound1);
        SafeArrayGetUBound(sa, 1, &uBound1);
        SafeArrayGetLBound(sa, 2, &lBound2);
        SafeArrayGetUBound(sa, 2, &uBound2);

        Value::Array arr;
        arr.resize(uBound1 - lBound1 + 1);

        for (LONG i = lBound1; i <= uBound1; i++)
        {
            arr[i - lBound1].resize(uBound2 - lBound2 + 1);

            for (LONG j = lBound2; j <= uBound2; j++)
            {
                VARIANT elem;
                VariantInit(&elem);

                LONG idx[2] = {i, j};
                SafeArrayGetElement(sa, idx, &elem);

                arr[i - lBound1][j - lBound2] = VariantConverter::fromVariant(elem);

                VariantClear(&elem);
            }
        }

        return Value(arr);
    }

    throw std::runtime_error("Unsupported VARIANT type");
}

VARIANT VariantConverter::toVariant(const Value &v)
{
    VARIANT ret;
    VariantInit(&ret);

    if (std::holds_alternative<std::monostate>(v.data_))
    {
        ret.vt = VT_EMPTY;
    }
    else if (std::holds_alternative<int>(v.data_))
    {
        ret.vt = VT_I4;
        ret.lVal = std::get<int>(v.data_);
    }
    else if (std::holds_alternative<double>(v.data_))
    {
        ret.vt = VT_R8;
        ret.dblVal = std::get<double>(v.data_);
    }
    else if (std::holds_alternative<bool>(v.data_))
    {
        ret.vt = VT_BOOL;
        ret.boolVal = std::get<bool>(v.data_) ? VARIANT_TRUE : VARIANT_FALSE;
    }
    else if (std::holds_alternative<std::wstring>(v.data_))
    {
        ret.vt = VT_BSTR;
        ret.bstrVal = SysAllocString(std::get<std::wstring>(v.data_).c_str());
    }
    else if (std::holds_alternative<DispatchWrapper>(v.data_))
    {
        ret.vt = VT_DISPATCH;
        ret.pdispVal = (IDispatch *)std::get<DispatchWrapper>(v.data_).raw();
        if (ret.pdispVal)
            ret.pdispVal->AddRef();
    }
    else if (std::holds_alternative<Value::Array>(v.data_))
    {
        const auto &arr = std::get<Value::Array>(v.data_);

        LONG rows = (LONG)arr.size();
        LONG cols = rows > 0 ? (LONG)arr[0].size() : 0;

        SAFEARRAYBOUND bounds[2];
        bounds[0].lLbound = 0;
        bounds[0].cElements = rows;
        bounds[1].lLbound = 0;
        bounds[1].cElements = cols;

        SAFEARRAY *sa = SafeArrayCreate(VT_VARIANT, 2, bounds);
        if (!sa)
            throw std::runtime_error("SafeArrayCreate failed");

        for (LONG i = 0; i < rows; i++)
        {
            for (LONG j = 0; j < cols; j++)
            {
                VARIANT elem = VariantConverter::toVariant(arr[i][j]);

                LONG idx[2] = {i, j};
                SafeArrayPutElement(sa, idx, &elem);

                VariantClear(&elem);
            }
        }

        ret.vt = VT_ARRAY | VT_VARIANT;
        ret.parray = sa;
    }

    return ret;
}

DispatchWrapper::DispatchWrapper(void *disp)
    : disp_(disp)
{
    if (disp_)
    {
        ((IDispatch *)disp_)->AddRef();
    }
}

DispatchWrapper::DispatchWrapper(const DispatchWrapper &other)
    : disp_(other.disp_)
{
    if (disp_)
    {
        ((IDispatch *)disp_)->AddRef();
    }
}

DispatchWrapper &DispatchWrapper::operator=(const DispatchWrapper &other)
{
    if (this != &other)
    {
        if (disp_)
        {
            ((IDispatch *)disp_)->Release();
        }
        disp_ = other.disp_;
        if (disp_)
        {
            ((IDispatch *)disp_)->AddRef();
        }
    }
    return *this;
}

DispatchWrapper::DispatchWrapper(DispatchWrapper &&other) noexcept
    : disp_(other.disp_)
{
    other.disp_ = nullptr;
}

DispatchWrapper &DispatchWrapper::operator=(DispatchWrapper &&other) noexcept
{
    if (this != &other)
    {
        if (disp_)
        {
            ((IDispatch *)disp_)->Release();
        }
        disp_ = other.disp_;
        other.disp_ = nullptr;
    }
    return *this;
}

DispatchWrapper::~DispatchWrapper()
{
    if (disp_)
    {
        ((IDispatch *)disp_)->Release();
    }
}

// 共通ヘルパ
static DISPID get_dispid(IDispatch *disp, const wchar_t *name)
{
    DISPID id;
    LPOLESTR n = (LPOLESTR)name;
    HRESULT hr = disp->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr))
    {
        throw std::runtime_error("GetIDsOfNames failed");
    }
    return id;
}

// VT_DISPATCH を返すプロパティ取得
DispatchWrapper DispatchWrapper::getDispatch(const wchar_t *name) const
{
    IDispatch *raw = (IDispatch *)disp_;
    if (!raw)
        return DispatchWrapper();

    DISPID id = get_dispid(raw, name);

    DISPPARAMS params = {nullptr, nullptr, 0, 0};
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = raw->Invoke(id, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                             DISPATCH_PROPERTYGET, &params,
                             &result, nullptr, nullptr);
    if (FAILED(hr))
    {
        VariantClear(&result);
        throw std::runtime_error("Invoke PROPERTYGET failed");
    }

    DispatchWrapper w;
    if (result.vt == VT_DISPATCH && result.pdispVal)
    {
        w = DispatchWrapper(result.pdispVal);
    }
    VariantClear(&result);
    return w;
}

DispatchWrapper DispatchWrapper::getDispatch(const wchar_t *name, const Value &arg) const
{
    IDispatch *raw = (IDispatch *)disp_;
    if (!raw)
        return DispatchWrapper();

    DISPID id = get_dispid(raw, name);

    VARIANT v = VariantConverter::toVariant(arg);

    DISPPARAMS params;
    params.cArgs = 1;
    params.rgvarg = &v;
    params.cNamedArgs = 0;
    params.rgdispidNamedArgs = nullptr;

    VARIANT result;
    VariantInit(&result);

    HRESULT hr = raw->Invoke(id, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                             DISPATCH_PROPERTYGET, &params,
                             &result, nullptr, nullptr);

    VariantClear(&v);

    if (FAILED(hr))
    {
        VariantClear(&result);
        throw std::runtime_error("Invoke PROPERTYGET (with arg) failed");
    }

    DispatchWrapper w;
    if (result.vt == VT_DISPATCH && result.pdispVal)
    {
        w = DispatchWrapper(result.pdispVal);
    }

    VariantClear(&result);
    return w;
}

// Value としてプロパティ取得
Value DispatchWrapper::getValue(const wchar_t *name) const
{
    IDispatch *raw = (IDispatch *)disp_;
    if (!raw)
        return Value();

    DISPID id = get_dispid(raw, name);

    DISPPARAMS params = {nullptr, nullptr, 0, 0};
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = raw->Invoke(id, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                             DISPATCH_PROPERTYGET, &params,
                             &result, nullptr, nullptr);
    if (FAILED(hr))
    {
        VariantClear(&result);
        throw std::runtime_error("Invoke PROPERTYGET failed");
    }

    Value v = VariantConverter::fromVariant(result);
    VariantClear(&result);
    return v;
}

// Value を使ったプロパティ設定
void DispatchWrapper::putValue(const wchar_t *name, const Value &val)
{
    IDispatch *raw = (IDispatch *)disp_;
    if (!raw)
        return;

    DISPID id = get_dispid(raw, name);

    VARIANT arg = VariantConverter::toVariant(val);

    DISPPARAMS params;
    params.cArgs = 1;
    params.rgvarg = &arg;
    params.cNamedArgs = 1;
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    params.rgdispidNamedArgs = &dispidNamed;

    HRESULT hr = raw->Invoke(id, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                             DISPATCH_PROPERTYPUT, &params,
                             nullptr, nullptr, nullptr);

    VariantClear(&arg);

    if (FAILED(hr))
    {
        throw std::runtime_error("Invoke PROPERTYPUT failed");
    }
}

// 引数なしメソッド呼び出し（VT_DISPATCH を返す想定）
DispatchWrapper DispatchWrapper::call(const wchar_t *name) const
{
    IDispatch *raw = (IDispatch *)disp_;
    if (!raw)
        return DispatchWrapper();

    DISPID id = get_dispid(raw, name);

    DISPPARAMS params = {nullptr, nullptr, 0, 0};
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = raw->Invoke(id, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                             DISPATCH_METHOD, &params,
                             &result, nullptr, nullptr);
    if (FAILED(hr))
    {
        VariantClear(&result);
        throw std::runtime_error("Invoke METHOD failed");
    }

    DispatchWrapper w;
    if (result.vt == VT_DISPATCH && result.pdispVal)
    {
        w = DispatchWrapper(result.pdispVal);
    }
    VariantClear(&result);
    return w;
}

Range::Range(const DispatchWrapper &disp)
    : disp_(disp)
{
}

void Range::setValue(const Value &v)
{
    disp_.putValue(L"Value", v);
}

Value Range::getValue() const
{
    return disp_.getValue(L"Value");
}

Worksheet::Worksheet(const DispatchWrapper &disp)
    : disp_(disp)
{
}

Range Worksheet::range(const wchar_t *ref)
{
    Value v(ref);
    auto r = disp_.getDispatch(L"Range", v);
    return Range(r);
}

void Worksheet::setName(const std::wstring &name)
{
    disp_.putValue(L"Name", Value(name));
}

Workbook::Workbook(const DispatchWrapper &disp)
    : disp_(disp)
{
}

Worksheet Workbook::activeSheet()
{
    auto sheet = disp_.getDispatch(L"ActiveSheet");
    return Worksheet(sheet);
}

void Workbook::save()
{
    disp_.call(L"Save");
}

void Workbook::saveAs(const std::wstring& path)
{
    // 相対パス → 絶対パス
    std::filesystem::path p(path);
    std::filesystem::path abs = std::filesystem::absolute(p);

    IDispatch* raw = (IDispatch*)disp_.raw();
    if (!raw)
        return;

    DISPID id;
    LPOLESTR name = (LPOLESTR)L"SaveAs";
    HRESULT hr = raw->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr))
        throw std::runtime_error("GetIDsOfNames(SaveAs) failed");

    VARIANT arg;
    VariantInit(&arg);
    arg.vt = VT_BSTR;
    arg.bstrVal = SysAllocString(abs.c_str());  // ★ 絶対パスを渡す

    DISPPARAMS params;
    params.cArgs = 1;
    params.rgvarg = &arg;
    params.cNamedArgs = 0;
    params.rgdispidNamedArgs = nullptr;

    hr = raw->Invoke(id, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                     DISPATCH_METHOD, &params,
                     nullptr, nullptr, nullptr);

    VariantClear(&arg);

    if (FAILED(hr))
        throw std::runtime_error("Invoke SaveAs failed");
}

void Workbook::close(bool saveChanges)
{
    disp_.call(L"Close");
}

Application::Application()
    : disp_(SingleCall::call())
{
}

void Application::visible(bool v)
{
    disp_.putValue(L"Visible", v ? 1 : 0);
}

void Application::setDisplayAlerts(bool v)
{
    disp_.putValue(L"DisplayAlerts", Value(v));
}

Workbook Application::addWorkbook()
{
    auto books = disp_.getDispatch(L"Workbooks");
    auto book = books.call(L"Add");
    return Workbook(book);
}

void Application::quit()
{
    disp_.call(L"Quit");
}
