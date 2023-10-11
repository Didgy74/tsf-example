cmake_minimum_required(VERSION 3.20)
project(tsf_test)

add_executable(tsf_test WIN32 main.cpp)
target_compile_features(tsf_test PUBLIC cxx_std_23)
set_property(TARGET tsf_test PROPERTY CXX_STANDARD 23)
set_property(TARGET tsf_test PROPERTY CXX_STANDARD_REQUIRED ON)