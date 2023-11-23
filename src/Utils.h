#pragma once

#include "Box.h"

#include "pugixml.hpp"

namespace dxfrg
{
	int Pt2Twip( const double pt );
	double Twip2Pt( const int twip );

	double Inch2Pt( const double inch );
	double Pt2Inch( const double pt );

	double MM2Inch( const int mm );
	int Inch2MM( const double inch );

	double CM2Inch( const double cm );
	double Inch2CM( const double inch );

	int Inch2Twip( const double inch );
	double Twip2Inch( const int twip );

	int MM2Twip( const int mm );
	int Twip2MM( const int twip );

	int CM2Twip( const double cm );
	double Twip2CM( const int twip );

	int CM2Emu( const double cm );

	pugi::xml_node GetLastChild( pugi::xml_node node, const char* name );
	void SetBorders( pugi::xml_node& w_bdrs, const char* elemName, const Box::BorderStyle style, const double width, const char* color );
}