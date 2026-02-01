#include <iostream>
#include "cpp_xlsx.hpp"

int main()
{
    try
    {
        cpp_xlsx::Application app;
        app.visible(true);
        app.setDisplayAlerts(false);

        cpp_xlsx::Workbook book = app.addWorkbook();
        cpp_xlsx::Worksheet sheet = book.activeSheet();

        sheet.setName(L"売上データ2025");
        sheet.range(L"A1:B2").setValue(cpp_xlsx::Value::Array{{cpp_xlsx::Value(1.55), cpp_xlsx::Value(2.333) }, {cpp_xlsx::Value(3.14), cpp_xlsx::Value(4)}});

        
        // std::wcout << arr[0][0].toString() << std::endl;
        // std::wcout << arr[0][1].toString() << std::endl;
        // std::wcout << arr[1][0].toString() << std::endl;
        // std::wcout << arr[1][1].toString() << std::endl;

        book.saveAs(L"test.xlsx");
        book.close(false);
        app.quit();
    }
    catch (const std::exception &e)
    {
        std::cerr << "Error: " << e.what() << std::endl;
    }

    return 0;
}
