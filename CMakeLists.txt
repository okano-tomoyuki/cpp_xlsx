cmake_minimum_required(VERSION 3.10)
project(cpp_xlsx VERSION 0.1 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${PROJECT_SOURCE_DIR})
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${PROJECT_SOURCE_DIR}/lib)
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${PROJECT_SOURCE_DIR}/lib)

add_executable(app
    main.cpp
)

add_library(${PROJECT_NAME}
    ${PROJECT_NAME}.cpp
)

target_link_libraries(${PROJECT_NAME}
    PRIVATE
        ole32
        oleaut32
        uuid
)

target_link_libraries(app
    PRIVATE
        ${PROJECT_NAME}
)