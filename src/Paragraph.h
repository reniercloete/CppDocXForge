#pragma once

#include "Box.h"
#include "Run.h"

#include "pugixml.hpp"

#include <iostream>

namespace dxfrg
{
    class Section;

    class Paragraph : public Box
    {
        friend class Document;
        friend class Section;
        friend class TableCell;
        friend std::ostream& operator<<( std::ostream& out, const Paragraph& p );

    public:
        // constructs an empty paragraph
        Paragraph();
        Paragraph( const Paragraph& p );
        ~Paragraph();
        void operator=( const Paragraph& right );
        //bool operator==( const Paragraph& p );
        friend bool operator==( const Paragraph& left, const Paragraph& right );

        operator bool();
        Paragraph Next();
        Paragraph Prev();

        // get run
        Run FirstRun();

        // add run
        Run AppendRun();
        Run AppendRun( const std::string& text );
        Run AppendRun( const std::string& text, const double fontSize );
        Run AppendRun( const std::string& text, const double fontSize, const std::string& fontAscii, const std::string& fontEastAsia = "" );
        Run AppendPageBreak();

        // paragraph formatting
        enum class Alignment { Left, Centered, Right, Justified, Distributed };
        void SetAlignment( const Alignment alignment );

        void SetLineSpacingSingle();               // Single
        void SetLineSpacingLines( const double at ); // 1.5 lines, Double (2 lines), Multiple (3 lines)
        void SetLineSpacingAtLeast( const int at );  // At Least
        void SetLineSpacingExactly( const int at );  // Exactly
        void SetLineSpacing( const int at, const char* lineRule );

        void SetBeforeSpacingAuto();
        void SetAfterSpacingAuto();
        void SetSpacingAuto( const char* attrNameAuto );
        void SetBeforeSpacingLines( const double beforeSpacing );
        void SetAfterSpacingLines( const double afterSpacing );
        void SetBeforeSpacing( const int beforeSpacing );
        void SetAfterSpacing( const int afterSpacing );
        void SetSpacing( const int twip, const char* attrNameAuto, const char* attrName );

        void SetLeftIndentChars( const double leftIndent );
        void SetRightIndentChars( const double rightIndent );
        void SetLeftIndent( const int leftIndent );
        void SetRightIndent( const int rightIndent );
        void SetFirstLineChars( const double indent );
        void SetHangingChars( const double indent );
        void SetFirstLine( const int indent );
        void SetHanging( const int indent );
        void SetIndent( const int indent, const char* attrName );

        void SetTopBorder( const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto" );
        void SetBottomBorder( const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto" );
        void SetLeftBorder( const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto" );
        void SetRightBorder( const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto" );
        void SetBorders( const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto" );
        void SetBorders_( const char* elemName, const BorderStyle style, const double width, const char* color );

        // helper
        void SetFontSize( const double fontSize );
        void SetFont( const std::string& fontAscii, const std::string& fontEastAsia = "" );
        void SetFontStyle( const Run::FontStyle fontStyle );
        void SetCharacterSpacing( const int characterSpacing );
        std::string GetText();

        // section
        Section GetSection();
        Section InsertSectionBreak();
        Section RemoveSectionBreak();
        bool HasSectionBreak();

    protected:
        struct Impl;
        Impl* impl_ = nullptr;

        // constructs paragraph from existing xml node
        Paragraph( Impl* impl );
    }; // class Paragraph

    struct Paragraph::Impl
    {
        pugi::xml_node w_body_;
        pugi::xml_node w_p_;
        pugi::xml_node w_pPr_;
    };

    bool dxfrg::operator==( const Paragraph& left, const Paragraph& right );
}