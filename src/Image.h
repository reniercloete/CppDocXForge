#pragma once

#include "Paragraph.h"

#include "pugixml.hpp"

namespace dxfrg
{
    class Image : public Paragraph
    {
        friend class Document;
    public:
        // constructs an empty text frame
        Image();
        Image( const Image& tf );
        ~Image();

       /* void SetSize( const int w, const int h );

        enum class Anchor { Page, Margin };
        enum class Position { Left, Center, Right, Top, Bottom };
        void SetAnchor_( const char* attrName, const Anchor anchor );
        void SetPosition_( const char* attrName, const Position align );
        void SetPosition_( const char* attrName, const int twip );

        void SetPositionX( const Position align, const Anchor ralativeTo );
        void SetPositionY( const Position align, const Anchor ralativeTo );
        void SetPositionX( const int x, const Anchor ralativeTo );
        void SetPositionY( const int y, const Anchor ralativeTo );

        enum class Wrapping { Around, None };
        void SetTextWrapping( const Wrapping wrapping );*/

    private:
        struct Impl;
        Impl* impl_ = nullptr;

        // constructs text frame from existing xml node
        Image( Impl* impl, Paragraph::Impl* p_impl );

    }; // class Image

    struct Image::Impl
    {
        //pugi::xml_node w_drawing_;
    };
}