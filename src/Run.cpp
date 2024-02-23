#include "Run.h"

#include "XMLWriter.h"

#include <sstream>
#include <iomanip>

using namespace dxfrg;

std::ostream& 
dxfrg::operator<<( std::ostream& out, const Run& r )
{
    if (r.impl_) {
        xml_string_writer writer;
        r.impl_->w_r_.print( writer, "  " );
        out << writer.Result;
    }
    else {
        out << "<run />";
    }
    return out;
}

// class Run
Run::Run()
{

}

Run::Run( Impl* impl ) : impl_( impl )
{

}

Run::Run( const Run& r )
{
    impl_ = new Impl;
    impl_->w_p_ = r.impl_->w_p_;
    impl_->w_r_ = r.impl_->w_r_;
    impl_->w_rPr_ = r.impl_->w_rPr_;
}

Run::~Run()
{
    if (impl_ != nullptr) {
        delete impl_;
        impl_ = nullptr;
    }
}

void Run::AppendText( const std::string& text )
{
    if (!impl_) return;
    auto t = impl_->w_r_.append_child( "w:t" );
    if (std::isspace( static_cast<unsigned char>(text.front()) ) || std::isspace( static_cast<unsigned char>(text.back()) )) {
        t.append_attribute( "xml:space" ) = "preserve";
    }
    t.text().set( text.c_str() );
}

std::string Run::GetText()
{
    if (!impl_) return "";
    std::string text;
    for (auto t = impl_->w_r_.child( "w:t" ); t; t = t.next_sibling( "w:t" )) {
        text += t.text().get();
    }
    return text;
}

void Run::ClearText()
{
    if (!impl_) return;
    impl_->w_r_.remove_children();
}

void Run::AppendLineBreak()
{
    if (!impl_) return;
    impl_->w_r_.append_child( "w:br" );
}

void 
Run::SetFontColor( const unsigned int FontColor )
{
    if (!impl_) return;
    auto sz = impl_->w_rPr_.child( "w:color" );
    if (!sz) {
        sz = impl_->w_rPr_.append_child( "w:color" );
    }
    auto szVal = sz.attribute( "w:val" );
    if (!szVal) {
        szVal = sz.append_attribute( "w:val" );
    }

    auto R = (FontColor & 0xFF0000) >> 16;
    auto G = (FontColor & 0x00FF00) >> 8;
    auto B = (FontColor & 0x0000FF);

    std::stringstream stream;
    stream 
        << std::hex << std::setfill( '0' ) 
        << std::setw( 2 ) << B 
        << std::setw( 2 ) << G 
        << std::setw( 2 ) << R;
    std::string result( stream.str() );

    szVal.set_value( result.c_str() );
}

unsigned int  
Run::GetFontCOLOR()
{
    if (!impl_) return -1;
    auto sz = impl_->w_rPr_.child( "w:color" );
    if (!sz) return 0;
    auto szVal = sz.attribute( "w:val" );
    if (!szVal) return 0;
    return szVal.as_uint();
}

void 
Run::SetFontSize( const double fontSize )
{
    if (!impl_) return;
    auto sz = impl_->w_rPr_.child( "w:sz" );
    if (!sz) {
        sz = impl_->w_rPr_.append_child( "w:sz" );
    }
    auto szVal = sz.attribute( "w:val" );
    if (!szVal) {
        szVal = sz.append_attribute( "w:val" );
    }
    // font size in half-points (1/144 of an inch)
    szVal.set_value( fontSize * 2 );
}

double Run::GetFontSize()
{
    if (!impl_) return -1;
    auto sz = impl_->w_rPr_.child( "w:sz" );
    if (!sz) return 0;
    auto szVal = sz.attribute( "w:val" );
    if (!szVal) return 0;
    return szVal.as_int() / 2.0;
}

void Run::SetFont(
    const std::string& fontAscii,
    const std::string& fontEastAsia )
{
    if (!impl_) return;
    auto rFonts = impl_->w_rPr_.child( "w:rFonts" );
    if (!rFonts) {
        rFonts = impl_->w_rPr_.append_child( "w:rFonts" );
    }
    auto rFontsAscii = rFonts.attribute( "w:ascii" );
    if (!rFontsAscii) {
        rFontsAscii = rFonts.append_attribute( "w:ascii" );
    }
    auto rFontsEastAsia = rFonts.attribute( "w:eastAsia" );
    if (!rFontsEastAsia) {
        rFontsEastAsia = rFonts.append_attribute( "w:eastAsia" );
    }
    rFontsAscii.set_value( fontAscii.c_str() );
    rFontsEastAsia.set_value( fontEastAsia.empty()
                              ? fontAscii.c_str()
                              : fontEastAsia.c_str() );
}

void Run::GetFont(
    std::string& fontAscii,
    std::string& fontEastAsia )
{
    if (!impl_) return;
    auto rFonts = impl_->w_rPr_.child( "w:rFonts" );
    if (!rFonts) return;

    auto rFontsAscii = rFonts.attribute( "w:ascii" );
    if (rFontsAscii) fontAscii = rFontsAscii.value();

    auto rFontsEastAsia = rFonts.attribute( "w:eastAsia" );
    if (rFontsEastAsia) fontEastAsia = rFontsEastAsia.value();
}

void 
Run::SetVerticalAlign( const VerticalAlign Value )
{
    if (!impl_) return;
    auto v = impl_->w_rPr_.child( "w:vertAlign" );

    if (Value & SubScript) {
        if (v.empty()) impl_->w_rPr_.append_child( "w:vertAlign" ).append_attribute( "w:val" ) = "subscript";
    }
    else {
        impl_->w_rPr_.remove_child( v );
    }

    if (Value & SuperScript) {
        if (v.empty()) impl_->w_rPr_.append_child( "w:vertAlign" ).append_attribute( "w:val" ) = "superscript";
    }
    else {
        impl_->w_rPr_.remove_child( v );
    }
}

Run::VerticalAlign 
Run::GetVerticalAlign()
{
    VerticalAlign Value = Baseline;
    if (!impl_) return Value;

    auto v = impl_->w_rPr_.child( "w:vertAlign" );
    if (v) {
		auto val = v.attribute( "w:val" );
		if (val) 
        {
			if (strcmp( val.value(), "subscript" ) == 0) 
                Value = SubScript;
			else if (strcmp( val.value(), "superscript" ) == 0) 
                Value = SuperScript;
		}
	}
    return Value;
}

void 
Run::SetFontStyle( const FontStyle f )
{
    if (!impl_) return;
    auto b = impl_->w_rPr_.child( "w:b" );
    if (f & Bold) {
        if (b.empty()) impl_->w_rPr_.append_child( "w:b" );
    }
    else {
        impl_->w_rPr_.remove_child( b );
    }

    auto i = impl_->w_rPr_.child( "w:i" );
    if (f & Italic) {
        if (i.empty()) impl_->w_rPr_.append_child( "w:i" );
    }
    else {
        impl_->w_rPr_.remove_child( i );
    }

    auto u = impl_->w_rPr_.child( "w:u" );
    if (f & Underline) {
        if (u.empty())
            impl_->w_rPr_.append_child( "w:u" ).append_attribute( "w:val" ) = "single";
    }
    else {
        impl_->w_rPr_.remove_child( u );
    }

    auto strike = impl_->w_rPr_.child( "w:strike" );
    if (f & Strikethrough) {
        if (strike.empty())
            impl_->w_rPr_.append_child( "w:strike" ).append_attribute( "w:val" ) = "true";
    }
    else {
        impl_->w_rPr_.remove_child( strike );
    }
}

Run::FontStyle Run::GetFontStyle()
{
    FontStyle fontStyle = 0;
    if (!impl_) return fontStyle;
    if (impl_->w_rPr_.child( "w:b" )) fontStyle |= Bold;
    if (impl_->w_rPr_.child( "w:i" )) fontStyle |= Italic;
    if (impl_->w_rPr_.child( "w:u" )) fontStyle |= Underline;
    if (impl_->w_rPr_.child( "w:strike" )) fontStyle |= Strikethrough;
    return fontStyle;
}

void Run::SetCharacterSpacing( const int characterSpacing )
{
    if (!impl_) return;
    auto spacing = impl_->w_rPr_.child( "w:spacing" );
    if (!spacing) {
        spacing = impl_->w_rPr_.append_child( "w:spacing" );
    }
    auto spacingVal = spacing.attribute( "w:val" );
    if (!spacingVal) {
        spacingVal = spacing.append_attribute( "w:val" );
    }
    spacingVal.set_value( characterSpacing );
}

int Run::GetCharacterSpacing()
{
    if (!impl_) return -1;
    return impl_->w_rPr_.child( "w:spacing" ).attribute( "w:val" ).as_int();
}

bool Run::IsPageBreak()
{
    if (!impl_) return false;
    return impl_->w_r_.find_child_by_attribute( "w:br", "w:type", "page" );
}

void Run::Remove()
{
    if (!impl_) return;
    impl_->w_p_.remove_child( impl_->w_r_ );
}

Run Run::Next()
{
    if (!impl_) return Run();
    auto w_r = impl_->w_r_.next_sibling( "w:r" );
    if (!w_r) return Run();

    Run::Impl* impl = new Run::Impl;
    impl->w_p_ = impl_->w_p_;
    impl->w_r_ = w_r;
    impl->w_rPr_ = w_r.child( "w:rPr" );
    return Run( impl );
}

Run::operator bool()
{
    return impl_ != nullptr && impl_->w_r_;
}

void Run::operator=( const Run& right )
{
    if (this == &right) return;
    if (impl_ != nullptr) delete impl_;
    if (right.impl_ != nullptr) {
        impl_ = new Impl;
        impl_->w_p_ = right.impl_->w_p_;
        impl_->w_r_ = right.impl_->w_r_;
        impl_->w_rPr_ = right.impl_->w_rPr_;
    }
    else {
        impl_ = nullptr;
    }
}