cmake_minimum_required(VERSION 3.15)
project(LookbackMCExcel LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

add_library(LookbackMC SHARED
    LookbackMC.cpp
)

target_include_directories(LookbackMC PUBLIC ${CMAKE_CURRENT_SOURCE_DIR})

# On Windows, produce LookbackMC.dll
