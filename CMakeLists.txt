cmake_minimum_required(VERSION 3.20)

project(gram_repro LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

add_library(gram_core
    src/spectrogram_core.cpp
    src/spectrogram_core.h
)

target_include_directories(gram_core PUBLIC ${CMAKE_CURRENT_SOURCE_DIR}/src)

add_executable(gram_repro
    src/cli_main.cpp
)

target_link_libraries(gram_repro PRIVATE gram_core)

add_executable(gram_live WIN32
    src/win_live_app.cpp
)

add_executable(gram_audio_info WIN32
    src/audio_info_app.cpp
)

target_link_libraries(gram_live PRIVATE gram_core winmm comdlg32 comctl32 runtimeobject)
target_link_libraries(gram_live PRIVATE ole32 uuid)
target_link_libraries(gram_audio_info PRIVATE comctl32 ole32 uuid)

if(MSVC)
    target_compile_options(gram_core PRIVATE /W4 /permissive-)
    target_compile_options(gram_repro PRIVATE /W4 /permissive-)
    target_compile_options(gram_live PRIVATE /W4 /permissive-)
    target_compile_options(gram_audio_info PRIVATE /W4 /permissive-)
else()
    target_compile_options(gram_core PRIVATE -Wall -Wextra -Wpedantic)
    target_compile_options(gram_repro PRIVATE -Wall -Wextra -Wpedantic)
    target_compile_options(gram_live PRIVATE -Wall -Wextra -Wpedantic)
    target_compile_options(gram_audio_info PRIVATE -Wall -Wextra -Wpedantic)
endif()
