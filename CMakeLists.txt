
cmake_minimum_required(VERSION 3.21)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

project(CppDocXForge VERSION 0.5.0 LANGUAGES C CXX) # C needed by zip.c

option(BUILD_EXAMPLES  "Build examples"               ON)
option(WITH_STATIC_CRT "Use static C Runtime library" OFF)

set(THIRD_PARTY_DIR "${CMAKE_CURRENT_SOURCE_DIR}/3rdparty" CACHE PATH "3rdparty")
set(ZIP_DIR         "${THIRD_PARTY_DIR}/zip-0.1.30"         CACHE PATH "zip")
set(PUGIXML_DIR     "${THIRD_PARTY_DIR}/pugixml-1.13"      CACHE PATH "pugixml")

set(CMAKE_CONFIGURATION_TYPES Debug Release)
set(CMAKE_DEBUG_POSTFIX d)

add_library(zip INTERFACE)
set_target_properties(zip PROPERTIES
  INTERFACE_INCLUDE_DIRECTORIES "${ZIP_DIR}"
  INTERFACE_SOURCES             "${ZIP_DIR}/zip.c"
)

add_library(pugixml INTERFACE)
set_target_properties(pugixml PROPERTIES
  INTERFACE_INCLUDE_DIRECTORIES "${PUGIXML_DIR}"
  INTERFACE_SOURCES             "${PUGIXML_DIR}/pugixml.cpp"
)

add_library(CppDocXForge STATIC)
list(APPEND sources 
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Document.cpp"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Examples.cpp"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Image.cpp"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Paragraph.cpp"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Run.cpp"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Section.cpp"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Table.cpp"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/TableCell.cpp"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/TextFrame.cpp"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Utils.cpp"
	)
list(APPEND headers 
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Box.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Cell.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Constants.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Document.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Image.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Paragraph.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Run.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Section.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Table.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/TableCell.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/TextFrame.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Utils.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/XMLWriter.h"
	)
source_group("Header Files" FILES ${headers})
target_sources(CppDocXForge PRIVATE ${sources} ${headers})
target_include_directories(CppDocXForge PUBLIC "${CMAKE_CURRENT_SOURCE_DIR}/src")
target_link_libraries(CppDocXForge PRIVATE zip pugixml)
if(WITH_STATIC_CRT)
  set_target_properties(CppDocXForge PROPERTIES MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")
endif()

if(BUILD_EXAMPLES)
  link_libraries(CppDocXForge)
  add_executable(examples src/Examples.cpp)
  target_include_directories(examples PUBLIC "${PUGIXML_DIR}")
endif()

install(TARGETS CppDocXForge)
install(FILES
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Box.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Cell.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Constants.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Document.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Image.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Paragraph.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Run.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Section.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Table.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/TableCell.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/TextFrame.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/Utils.h"
		"${CMAKE_CURRENT_SOURCE_DIR}/src/XMLWriter.h"
	TYPE INCLUDE)





