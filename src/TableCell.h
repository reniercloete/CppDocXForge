#pragma once

#include "Cell.h"

#include "pugixml.hpp"

namespace dxfrg
{
    class Paragraph;

    class TableCell
    {
        friend class Table;

    public:
        // constructs an empty cell
        TableCell();
        TableCell( const TableCell& tc );
        ~TableCell();
        void operator=( const TableCell& right );

        operator bool();
        bool empty() const;

        void SetWidth( const int w, const char* units = "dxa" );
        void SetHeight( const int h, const char* units = "exact" );

        enum class Alignment { Top, Center, Bottom };
        void SetVerticalAlignment( const Alignment align );

        void SetCellSpanning_( const int cols );

        Paragraph AppendParagraph();
        Paragraph FirstParagraph();

    private:
        struct Impl;
        Impl* impl_ = nullptr;

        // constructs a table from existing xml node
        TableCell( Impl* impl );
    }; // class TableCell

    struct TableCell::Impl
    {
        Cell* c_;
        pugi::xml_node w_tr_;
        pugi::xml_node w_tc_;
        pugi::xml_node w_tcPr_;
        pugi::xml_node w_trPr_;
    };
}