cmake_minimum_required(VERSION 3.20)

project(audio_spectrogram LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

add_library(gram_core
    src/spectrogram_core.cpp
    src/spectrogram_core.h
)

target_include_directories(gram_core PUBLIC ${CMAKE_CURRENT_SOURCE_DIR}/src)

add_executable(audio_spectrogram_cli
    src/cli_main.cpp
)

set_target_properties(audio_spectrogram_cli PROPERTIES OUTPUT_NAME "audio-spectrogram-cli")

target_link_libraries(audio_spectrogram_cli PRIVATE gram_core)

include(CTest)
if(BUILD_TESTING)
    add_executable(gram_core_tests tests/spectrogram_core_tests.cpp)
    target_link_libraries(gram_core_tests PRIVATE gram_core)
    add_test(NAME gram_core_tests COMMAND gram_core_tests)
    if(MSVC)
        target_compile_options(gram_core_tests PRIVATE /W4 /permissive-)
    else()
        target_compile_options(gram_core_tests PRIVATE -Wall -Wextra -Wpedantic)
    endif()
endif()

if(WIN32)
    add_executable(audio_spectrogram WIN32
        src/win_live_app.cpp
    )

    set_target_properties(audio_spectrogram PROPERTIES OUTPUT_NAME "audio-spectrogram")

    add_executable(audio_spectrogram_info WIN32
        src/audio_info_app.cpp
    )

    set_target_properties(audio_spectrogram_info PROPERTIES OUTPUT_NAME "audio-spectrogram-info")

    target_link_libraries(audio_spectrogram PRIVATE gram_core winmm comdlg32 comctl32 runtimeobject ole32 uuid)
    target_link_libraries(audio_spectrogram_info PRIVATE comctl32 ole32 uuid)
endif()

# Linux desktop application (Qt 6 Widgets). Optional: skipped silently when Qt 6 is not installed, so the CLI,
# the core and the tests build everywhere exactly as before.
if(UNIX AND NOT APPLE)
    option(AUDIO_SPECTROGRAM_BUILD_QT "Build the Linux Qt desktop application" ON)
    if(AUDIO_SPECTROGRAM_BUILD_QT)
        find_package(Qt6 QUIET COMPONENTS Widgets)
        if(Qt6_FOUND)
            add_library(gram_qt STATIC
                src/linux/pcm_capture.cpp
                src/linux/pcm_capture.h
                src/linux/spectrogram_view.cpp
                src/linux/spectrogram_view.h
                src/linux/main_window.cpp
                src/linux/main_window.h
            )
            set_target_properties(gram_qt PROPERTIES AUTOMOC ON)
            target_include_directories(gram_qt PUBLIC ${CMAKE_CURRENT_SOURCE_DIR}/src/linux)
            target_link_libraries(gram_qt PUBLIC gram_core Qt6::Widgets)

            add_executable(audio_spectrogram_qt src/linux/main.cpp)
            set_target_properties(audio_spectrogram_qt PROPERTIES OUTPUT_NAME "audio-spectrogram")
            target_link_libraries(audio_spectrogram_qt PRIVATE gram_qt)

            if(BUILD_TESTING)
                add_executable(gram_qt_tests tests/qt_gui_tests.cpp)
                target_link_libraries(gram_qt_tests PRIVATE gram_qt)
                add_test(NAME gram_qt_tests COMMAND gram_qt_tests)
                set_tests_properties(gram_qt_tests PROPERTIES ENVIRONMENT "QT_QPA_PLATFORM=offscreen")
            endif()

            include(GNUInstallDirs)
            install(TARGETS audio_spectrogram_qt audio_spectrogram_cli RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR})
            install(FILES packaging/linux/audio-spectrogram.desktop DESTINATION ${CMAKE_INSTALL_DATADIR}/applications)
            install(FILES packaging/linux/audio-spectrogram.svg
                    DESTINATION ${CMAKE_INSTALL_DATADIR}/icons/hicolor/scalable/apps)
        else()
            message(STATUS "Qt 6 Widgets not found: skipping the Linux desktop application")
        endif()
    endif()
endif()


if(MSVC)
    target_compile_options(gram_core PRIVATE /W4 /permissive-)
    target_compile_options(audio_spectrogram_cli PRIVATE /W4 /permissive-)
    if(WIN32)
        target_compile_options(audio_spectrogram PRIVATE /W4 /permissive-)
        target_compile_options(audio_spectrogram_info PRIVATE /W4 /permissive-)
    endif()
else()
    target_compile_options(gram_core PRIVATE -Wall -Wextra -Wpedantic)
    target_compile_options(audio_spectrogram_cli PRIVATE -Wall -Wextra -Wpedantic)
endif()
