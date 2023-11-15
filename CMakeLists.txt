cmake_minimum_required(VERSION 3.20)
project(tsf_example)

add_executable(tsf_example WIN32 main.cpp)
target_compile_features(tsf_example PUBLIC cxx_std_23)
set_property(TARGET tsf_example PROPERTY CXX_STANDARD 23)
set_property(TARGET tsf_example PROPERTY CXX_STANDARD_REQUIRED ON)