#pragma once

#include "pugixml.hpp"

#include <iostream>

namespace dxfrg
{
    class Paragraph;

    class Section
    {
        friend class Document;
        friend class Paragraph;
        friend std::ostream& operator<<( std::ostream& out, const Section& s );

    public:
        // constructs an empty section
        Section();
        Section( const Section& s );
        ~Section();
        void operator=( const Section& right );

        friend bool operator==( const Section& left, const Section& right );
        //bool operator==( const Section& s );

        operator bool();
        Section Next();
        Section Prev();

        // section
        void Split();
        bool IsSplit();
        void Merge();

        // page formatting
        enum class Orientation { Landscape, Portrait, Unknown };
        void SetPageSize( const int w, const int h );
        void GetPageSize( int& w, int& h );

        void SetPageOrient( const Orientation orient );
        Orientation GetPageOrient();

        void SetPageMargin( const int top, const int bottom, const int left, const int right );
        void GetPageMargin( int& top, int& bottom, int& left, int& right );

        void SetPageMargin( const int header, const int footer );
        void GetPageMargin( int& header, int& footer );

        void SetColumn( const int num, const int space = 425 );

        enum class PageNumberFormat {
            Decimal,      // e.g., 1, 2, 3, 4, etc.
            NumberInDash, // e.g., -1-, -2-, -3-, -4-, etc.
            CardinalText, // In English, One, Two, Three, etc.
            OrdinalText,  // In English, First, Second, Third, etc.
            LowerLetter,  // e.g., a, b, c, etc.
            UpperLetter,  // e.g., A, B, C, etc.
            LowerRoman,   // e.g., i, ii, iii, iv, etc.
            UpperRoman    // e.g., I, II, III, IV, etc.
        };

        /**
         * Specifies the page numbers for pages in the section.
         *
         * @param fmt   Specifies the number format to be used for page numbers in the section.
         *
         * @param start Specifies the page number that appears on the first page of the section.
         *              If the value is omitted, numbering continues from the highest page number in the previous section.
         */
        void SetPageNumber( const PageNumberFormat fmt = PageNumberFormat::Decimal, const unsigned int start = 0 );
        void RemovePageNumber();

        // paragraph
        Paragraph FirstParagraph();
        Paragraph LastParagraph();

    private:
        struct Impl;
        Impl* impl_ = nullptr;

        // constructs section from existing xml node
        Section( Impl* impl );
        void FindSectionProperties();
    }; // class Section

    struct Section::Impl
    {
        pugi::xml_node w_body_;
        pugi::xml_node w_p_;      // current paragraph
        pugi::xml_node w_pPr_;
        pugi::xml_node w_p_last_; // the last paragraph of the section
        pugi::xml_node w_pPr_last_;
        pugi::xml_node w_sectPr_;
    };

    bool operator==( const Section& left, const Section& right );
}
