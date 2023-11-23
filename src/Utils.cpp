#include "Utils.h"

namespace dxfrg
{
    int Pt2Twip( const double pt )
    {
        return pt * 20;
    }

    double Twip2Pt( const int twip )
    {
        return twip / 20.0;
    }

    double Inch2Pt( const double inch )
    {
        return inch * 72;
    }

    double Pt2Inch( const double pt )
    {
        return pt / 72;
    }

    double MM2Inch( const int mm )
    {
        return mm / 25.4;
    }

    int Inch2MM( const double inch )
    {
        return inch * 25.4;
    }

    double CM2Inch( const double cm )
    {
        return cm / 2.54;
    }

    double Inch2CM( const double inch )
    {
        return inch * 2.54;
    }

    int Inch2Twip( const double inch )
    {
        return inch * 1440;
    }

    double Twip2Inch( const int twip )
    {
        return twip / 1440.0;
    }

    int MM2Twip( const int mm )
    {
        return Inch2Twip( MM2Inch( mm ) );
    }

    int Twip2MM( const int twip )
    {
        return Inch2MM( Twip2Inch( twip ) );
    }

    int CM2Twip( const double cm )
    {
        return Inch2Twip( CM2Inch( cm ) );
    }

    int CM2Emu( const double cm )
    {
        return cm * 360000;
    }

    double Twip2CM( const int twip )
    {
        return Inch2CM( Twip2Inch( twip ) );
    }

    pugi::xml_node GetLastChild( pugi::xml_node node, const char* name )
    {
        pugi::xml_node child = node.last_child();
        while (!child.empty() && std::strcmp( name, child.name() ) != 0) {
            child = child.previous_sibling( name );
        }
        return child;
    }

    void SetBorders( pugi::xml_node& w_bdrs, const char* elemName, const Box::BorderStyle style, const double width, const char* color )
    {
        auto w_bdr = w_bdrs.child( elemName );
        if (!w_bdr) {
            w_bdr = w_bdrs.append_child( elemName );
        }

        const char* val = "";
        switch (style) {
            case Box::BorderStyle::Single:
                val = "single";
                break;
            case Box::BorderStyle::Dotted:
                val = "dotted";
                break;
            case Box::BorderStyle::DotDash:
                val = "dotDash";
                break;
            case Box::BorderStyle::Dashed:
                val = "dashed";
                break;
            case Box::BorderStyle::Double:
                val = "double";
                break;
            case Box::BorderStyle::Wave:
                val = "wave";
                break;
            case Box::BorderStyle::None:
                val = "none";
                break;
        }

        auto w_val = w_bdr.attribute( "w:val" );
        if (!w_val) {
            w_val = w_bdr.append_attribute( "w:val" );
        }
        w_val.set_value( val );

        auto w_sz = w_bdr.attribute( "w:sz" );
        if (!w_sz) {
            w_sz = w_bdr.append_attribute( "w:sz" );
        }
        w_sz.set_value( width * 8 );

        auto w_color = w_bdr.attribute( "w:color" );
        if (!w_color) {
            w_color = w_bdr.append_attribute( "w:color" );
        }
        w_color.set_value( color );
    }
}

