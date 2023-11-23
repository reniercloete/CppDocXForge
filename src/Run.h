#pragma once

#include "pugixml.hpp"

#include <iostream>

namespace dxfrg
{
    class Run
    {
        friend class Paragraph;
        friend std::ostream& operator<<( std::ostream& out, const Run& r );

    public:
        // constructs an empty run
        Run();
        Run( const Run& r );
        ~Run();
        void operator=( const Run& right );

        operator bool();
        Run Next();

        // text
        void AppendText( const std::string& text );
        std::string GetText();
        void ClearText();
        void AppendLineBreak();

        // text formatting
        using FontStyle = unsigned int;
        enum : FontStyle
        {
            Bold = 1,
            Italic = 2,
            Underline = 4,
            Strikethrough = 8
        };
        void SetFontSize( const double fontSize );
        double GetFontSize();

        void SetFont( const std::string& fontAscii, const std::string& fontEastAsia = "" );
        void GetFont( std::string& fontAscii, std::string& fontEastAsia );

        void SetFontStyle( const FontStyle fontStyle );
        FontStyle GetFontStyle();

        void SetCharacterSpacing( const int characterSpacing );
        int GetCharacterSpacing();

        // Run
        void Remove();
        bool IsPageBreak();

    private:
        struct Impl;
        Impl* impl_ = nullptr;

        // constructs run from existing xml node
        Run( Impl* impl );
    }; // class Run

    struct Run::Impl
    {
        pugi::xml_node w_p_;
        pugi::xml_node w_r_;
        pugi::xml_node w_rPr_;
    };
}