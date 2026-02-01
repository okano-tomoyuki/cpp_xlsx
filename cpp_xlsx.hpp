#ifndef CPP_XLSX_HPP
#define CPP_XLSX_HPP

#include <memory>
#include <string>
#include <stdexcept>
#include <variant>
#include <vector>

namespace cpp_xlsx
{

class Value;

class DispatchWrapper
{
public:
    DispatchWrapper() = default;
    explicit DispatchWrapper(void *disp);
    DispatchWrapper(const DispatchWrapper &other);
    DispatchWrapper &operator=(const DispatchWrapper &other);
    DispatchWrapper(DispatchWrapper &&other) noexcept;
    DispatchWrapper &operator=(DispatchWrapper &&other) noexcept;
    ~DispatchWrapper();
    DispatchWrapper getDispatch(const wchar_t *name) const;
    DispatchWrapper getDispatch(const wchar_t* name, const Value& arg) const;
    Value getValue(const wchar_t *name) const;
    void putValue(const wchar_t *name, const Value &v);
    DispatchWrapper call(const wchar_t *name) const;
    void *raw() const { return disp_; }

private:
    void *disp_ = nullptr;
};

class Value
{
public:
    using Array = std::vector<std::vector<Value>>;
    using VariantType = std::variant<
        std::monostate,  // Empty
        int,             // VT_I4
        double,          // VT_R8
        bool,            // VT_BOOL
        std::wstring,    // VT_BSTR
        DispatchWrapper, // VT_DISPATCH
        Array            // SAFEARRAY(VARIANT)
        >;

    Value() = default;
    Value(int v);
    Value(double v);
    Value(bool v);
    Value(const wchar_t *s);
    Value(const std::wstring &s);
    Value(const DispatchWrapper &d);
    Value(const Array &arr);
    const VariantType &data() const { return data_; }

    std::wstring toString() const;

private:
    VariantType data_;
    friend class VariantConverter;
};

class Range
{
public:
    explicit Range(const DispatchWrapper &disp);

    void setValue(const Value &v);
    Value getValue() const;

private:
    DispatchWrapper disp_;
};

class Worksheet
{
public:
    explicit Worksheet(const DispatchWrapper &disp);

    Range range(const wchar_t *ref);
    void setName(const std::wstring& name);

private:
    DispatchWrapper disp_;
};

class Workbook 
{
public:
    explicit Workbook(const DispatchWrapper& disp);

    Worksheet activeSheet();

    void save();
    void saveAs(const std::wstring& path);
    void close(bool saveChanges = false);

private:
    DispatchWrapper disp_;
};


class Application
{
public:
    Application();

    void visible(bool v);

    void setDisplayAlerts(bool v);

    void quit();

    Workbook addWorkbook();

private:
    DispatchWrapper disp_;
};

}

#endif