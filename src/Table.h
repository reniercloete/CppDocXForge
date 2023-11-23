#pragma once

#include "Box.h"
#include "Cell.h"

#include "pugixml.hpp"

#include <iostream>
#include <vector>

namespace dxfrg
{
    class TableCell;

    class Table : public Box
    {
        friend class Document;

    public:
        // constructs an empty table
        Table();
        Table( const Table& t );
        ~Table();
        void operator=( const Table& right );

        void Create_( const int rows, const int cols );

        TableCell GetCell( const int row, const int col );
        TableCell GetCell_( const int row, const int col );
        bool MergeCells( TableCell tc1, TableCell tc2 );
        bool SplitCell();

        void RemoveCell_( TableCell tc );

        // units: 
        //   auto - Specifies that width is determined by the overall table layout algorithm.
        //   dxa  - Specifies that the value is in twentieths of a point (1/1440 of an inch).
        //   nil  - Specifies a value of zero.
        //   pct  - Specifies a value as a percent of the table width.
        void SetWidthAuto();
        void SetWidthPercent( const double w ); // 0-100
        void SetWidth( const int w, const char* units = "dxa" );

        // the distance between the cell contents and the cell borders
        void SetCellMarginTop( const int w, const char* units = "dxa" );
        void SetCellMarginBottom( const int w, const char* units = "dxa" );
        void SetCellMarginLeft( const int w, const char* units = "dxa" );
        void SetCellMarginRight( const int w, const char* units = "dxa" );
        void SetCellMargin( const char* elemName, const int w, const char* units = "dxa" );

        // table formatting
        enum class Alignment { Left, Centered, Right };
        void SetAlignment( const Alignment alignment );

        // style - Specifies the style of the border.
        // width - Specifies the width of the border in points.
        // color - Specifies the color of the border. 
        //         Values are given as hex values (in RRGGBB format). 
        //         No #, unlike hex values in HTML/CSS. E.g., color="FFFF00". 
        //         A value of auto is also permitted and will allow the 
        //         consuming word processor to determine the color.
        void SetTopBorders( const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto" );
        void SetBottomBorders( const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto" );
        void SetLeftBorders( const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto" );
        void SetRightBorders( const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto" );
        void SetInsideHBorders( const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto" );
        void SetInsideVBorders( const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto" );
        void SetInsideBorders( const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto" );
        void SetOutsideBorders( const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto" );
        void SetAllBorders( const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto" );
        void SetBorders_( const char* elemName, const BorderStyle style, const double width, const char* color );

    private:
        struct Impl;
        Impl* impl_ = nullptr;

        // constructs a table from existing xml node
        Table( Impl* impl );
    }; // class Table

    using Row = std::vector<Cell>;
    using Grid = std::vector<Row>;

    struct Table::Impl
    {
        pugi::xml_node w_body_;
        pugi::xml_node w_tbl_;
        pugi::xml_node w_tblPr_;
        pugi::xml_node w_tblGrid_;

        int rows_ = 0;
        int cols_ = 0;
        Grid grid_; // logical grid
    };
}