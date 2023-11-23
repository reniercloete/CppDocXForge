#include "Paragraph.h"

#include "Utils.h"
#include "Section.h"

#include "XMLWriter.h"

using namespace dxfrg;

std::ostream& 
dxfrg::operator<<( std::ostream& out, const Paragraph& p )
{
    if (p.impl_) {
        xml_string_writer writer;
        p.impl_->w_p_.print( writer, "  " );
        out << writer.Result;
    }
    else 
    {
        out << "<paragraph />";
    }
    return out;
}

// class Paragraph
Paragraph::Paragraph()
{

}

Paragraph::Paragraph( Impl* impl ) : impl_( impl )
{

}

Paragraph::Paragraph( const Paragraph& p )
{
    impl_ = new Impl;
    impl_->w_body_ = p.impl_->w_body_;
    impl_->w_p_ = p.impl_->w_p_;
    impl_->w_pPr_ = p.impl_->w_pPr_;
}

Paragraph::~Paragraph()
{
    if (impl_ != nullptr) {
        delete impl_;
        impl_ = nullptr;
    }
}

Run Paragraph::FirstRun()
{
    if (!impl_) return Run();
    auto w_r = impl_->w_p_.child( "w:r" );
    if (!w_r) return Run();

    Run::Impl* impl = new Run::Impl;
    impl->w_p_ = impl_->w_p_;
    impl->w_r_ = w_r;
    impl->w_rPr_ = w_r.child( "w:rPr" );
    return Run( impl );
}

Run Paragraph::AppendRun()
{
    if (!impl_) return Run();
    auto w_r = impl_->w_p_.append_child( "w:r" );
    auto w_rPr = w_r.append_child( "w:rPr" );

    Run::Impl* impl = new Run::Impl;
    impl->w_p_ = impl_->w_p_;
    impl->w_r_ = w_r;
    impl->w_rPr_ = w_rPr;
    return Run( impl );
}

Run Paragraph::AppendRun( const std::string& text )
{
    auto r = AppendRun();
    if (!text.empty()) {
        r.AppendText( text );
    }
    return r;
}

Run Paragraph::AppendRun( const std::string& text, const double fontSize )
{
    auto r = AppendRun( text );
    if (fontSize != 0) {
        r.SetFontSize( fontSize );
    }
    return r;
}

Run Paragraph::AppendRun( const std::string& text, const double fontSize, const std::string& fontAscii, const std::string& fontEastAsia )
{
    auto r = AppendRun( text, fontSize );
    if (!fontAscii.empty()) {
        r.SetFont( fontAscii, fontEastAsia );
    }
    return r;
}

Run Paragraph::AppendPageBreak()
{
    if (!impl_) return Run();
    auto w_r = impl_->w_p_.append_child( "w:r" );
    auto w_br = w_r.append_child( "w:br" );
    w_br.append_attribute( "w:type" ) = "page";

    Run::Impl* impl = new Run::Impl;
    impl->w_p_ = impl_->w_p_;
    impl->w_r_ = w_r;
    impl->w_rPr_ = w_br;
    return Run( impl );
}

void Paragraph::SetAlignment( const Alignment alignment )
{
    if (!impl_) return;

    const char* val = "";
    switch (alignment) {
        case Alignment::Left:
            val = "start";
            break;
        case Alignment::Right:
            val = "end";
            break;
        case Alignment::Centered:
            val = "center";
            break;
        case Alignment::Justified:
            val = "both";
            break;
        case Alignment::Distributed:
            val = "distribute";
            break;
    }

    auto jc = impl_->w_pPr_.child( "w:jc" );
    if (!jc) {
        jc = impl_->w_pPr_.append_child( "w:jc" );
    }
    auto jcVal = jc.attribute( "w:val" );
    if (!jcVal) {
        jcVal = jc.append_attribute( "w:val" );
    }
    jcVal.set_value( val );
}

void Paragraph::SetLineSpacingSingle()
{
    if (!impl_) return;
    auto spacing = impl_->w_pPr_.child( "w:spacing" );
    if (!spacing) return;
    auto spacingLineRule = spacing.attribute( "w:lineRule" );
    if (spacingLineRule) {
        spacing.remove_attribute( spacingLineRule );
    }
    auto spacingLine = spacing.attribute( "w:line" );
    if (spacingLine) {
        spacing.remove_attribute( spacingLine );
    }
}

void Paragraph::SetLineSpacingLines( const double at )
{
    // A normal single-spaced paragaph has a w:line value of 240, or 12 points.
    // 
    // If the value of lineRule is auto, then the value of line 
    // is interpreted as 240th of a line, e.g. 360 = 1.5 lines.
    SetLineSpacing( at * 240, "auto" );
}

void Paragraph::SetLineSpacingAtLeast( const int at )
{
    // If the value of the lineRule attribute is atLeast or exactly, 
    // then the value of the line attribute is interpreted as 240th of a point.
    // (Not really. Actually, values are in twentieths of a point, e.g. 240 = 12 pt.)
    SetLineSpacing( at, "atLeast" );
}

void Paragraph::SetLineSpacingExactly( const int at )
{
    SetLineSpacing( at, "exact" );
}

void Paragraph::SetLineSpacing( const int at, const char* lineRule )
{
    if (!impl_) return;
    auto spacing = impl_->w_pPr_.child( "w:spacing" );
    if (!spacing) {
        spacing = impl_->w_pPr_.append_child( "w:spacing" );
    }

    auto spacingLineRule = spacing.attribute( "w:lineRule" );
    if (!spacingLineRule) {
        spacingLineRule = spacing.append_attribute( "w:lineRule" );
    }

    auto spacingLine = spacing.attribute( "w:line" );
    if (!spacingLine) {
        spacingLine = spacing.append_attribute( "w:line" );
    }

    spacingLineRule.set_value( lineRule );
    spacingLine.set_value( at );
}

void Paragraph::SetBeforeSpacingAuto()
{
    SetSpacingAuto( "w:beforeAutospacing" );
}

void Paragraph::SetAfterSpacingAuto()
{
    SetSpacingAuto( "w:afterAutospacing" );
}

void Paragraph::SetSpacingAuto( const char* attrNameAuto )
{
    if (!impl_) return;
    auto spacing = impl_->w_pPr_.child( "w:spacing" );
    if (!spacing) {
        spacing = impl_->w_pPr_.append_child( "w:spacing" );
    }
    auto spacingAuto = spacing.attribute( attrNameAuto );
    if (!spacingAuto) {
        spacingAuto = spacing.append_attribute( attrNameAuto );
    }
    // Any value for before or beforeLines is ignored.
    spacingAuto.set_value( "true" );
}

void Paragraph::SetBeforeSpacingLines( const double beforeSpacing )
{
    // To specify units in hundreths of a line, 
    // use attributes 'afterLines'/'beforeLines'.
    SetSpacing( beforeSpacing * 100, "w:beforeAutospacing", "w:beforeLines" );
}

void Paragraph::SetAfterSpacingLines( const double afterSpacing )
{
    SetSpacing( afterSpacing * 100, "w:afterAutospacing", "w:afterLines" );
}

void Paragraph::SetBeforeSpacing( const int beforeSpacing )
{
    SetSpacing( beforeSpacing, "w:beforeAutospacing", "w:before" );
}

void Paragraph::SetAfterSpacing( const int afterSpacing )
{
    SetSpacing( afterSpacing, "w:afterAutospacing", "w:after" );
}

void Paragraph::SetSpacing( const int twip, const char* attrNameAuto, const char* attrName )
{
    if (!impl_) return;
    auto elemSpacing = impl_->w_pPr_.child( "w:spacing" );
    if (!elemSpacing) {
        elemSpacing = impl_->w_pPr_.append_child( "w:spacing" );
    }

    auto attrSpacingAuto = elemSpacing.attribute( attrNameAuto );
    if (attrSpacingAuto) {
        elemSpacing.remove_attribute( attrSpacingAuto );
    }

    auto attrSpacing = elemSpacing.attribute( attrName );
    if (!attrSpacing) {
        attrSpacing = elemSpacing.append_attribute( attrName );
    }
    attrSpacing.set_value( twip );
}

void Paragraph::SetLeftIndentChars( const double leftIndent )
{
    // To specify units in hundreths of a character, 
    // use attributes leftChars/endChars, rightChars/endChars, etc. 
    SetIndent( leftIndent * 100, "w:leftChars" );
}

void Paragraph::SetRightIndentChars( const double rightIndent )
{
    SetIndent( rightIndent * 100, "w:rightChars" );
}

void Paragraph::SetLeftIndent( const int leftIndent )
{
    SetIndent( leftIndent, "w:left" );
}

void Paragraph::SetRightIndent( const int rightIndent )
{
    SetIndent( rightIndent, "w:right" );
}

void Paragraph::SetFirstLineChars( const double indent )
{
    SetIndent( indent * 100, "w:firstLineChars" );
}

void Paragraph::SetHangingChars( const double indent )
{
    SetIndent( indent * 100, "w:hangingChars" );
}

void Paragraph::SetFirstLine( const int indent )
{
    SetIndent( indent, "w:firstLine" );
}

void Paragraph::SetHanging( const int indent )
{
    SetIndent( indent, "w:hanging" );
    SetLeftIndent( indent );
}

void Paragraph::SetIndent( const int indent, const char* attrName )
{
    if (!impl_) return;
    auto elemIndent = impl_->w_pPr_.child( "w:ind" );
    if (!elemIndent) {
        elemIndent = impl_->w_pPr_.append_child( "w:ind" );
    }

    auto attrIndent = elemIndent.attribute( attrName );
    if (!attrIndent) {
        attrIndent = elemIndent.append_attribute( attrName );
    }
    attrIndent.set_value( indent );
}

void Paragraph::SetTopBorder( const BorderStyle style, const double width, const char* color )
{
    SetBorders_( "w:top", style, width, color );
}

void Paragraph::SetBottomBorder( const BorderStyle style, const double width, const char* color )
{
    SetBorders_( "w:bottom", style, width, color );
}

void Paragraph::SetLeftBorder( const BorderStyle style, const double width, const char* color )
{
    SetBorders_( "w:left", style, width, color );
}

void Paragraph::SetRightBorder( const BorderStyle style, const double width, const char* color )
{
    SetBorders_( "w:right", style, width, color );
}

void Paragraph::SetBorders( const BorderStyle style, const double width, const char* color )
{
    SetTopBorder( style, width, color );
    SetBottomBorder( style, width, color );
    SetLeftBorder( style, width, color );
    SetRightBorder( style, width, color );
}

void Paragraph::SetBorders_( const char* elemName, const BorderStyle style, const double width, const char* color )
{
    if (!impl_) return;
    auto w_pBdr = impl_->w_pPr_.child( "w:pBdr" );
    if (!w_pBdr) {
        w_pBdr = impl_->w_pPr_.append_child( "w:pBdr" );
    }
    dxfrg::SetBorders( w_pBdr, elemName, style, width, color );
}

void Paragraph::SetFontSize( const double fontSize )
{
    for (auto r = FirstRun(); r; r = r.Next()) {
        r.SetFontSize( fontSize );
    }
}

void Paragraph::SetFont( const std::string& fontAscii, const std::string& fontEastAsia )
{
    for (auto r = FirstRun(); r; r = r.Next()) {
        r.SetFont( fontAscii, fontEastAsia );
    }
}

void Paragraph::SetFontStyle( const Run::FontStyle fontStyle )
{
    for (auto r = FirstRun(); r; r = r.Next()) {
        r.SetFontStyle( fontStyle );
    }
}

void Paragraph::SetCharacterSpacing( const int characterSpacing )
{
    for (auto r = FirstRun(); r; r = r.Next()) {
        r.SetCharacterSpacing( characterSpacing );
    }
}

std::string Paragraph::GetText()
{
    std::string text;
    for (auto r = FirstRun(); r; r = r.Next()) {
        text += r.GetText();
    }
    return text;
}

Paragraph Paragraph::Next()
{
    if (!impl_) return Paragraph();
    auto w_p = impl_->w_p_.next_sibling( "w:p" );
    if (!w_p) return Paragraph();

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_p.child( "w:pPr" );
    return Paragraph( impl );
}

Paragraph Paragraph::Prev()
{
    if (!impl_) return Paragraph();
    auto w_p = impl_->w_p_.previous_sibling( "w:p" );
    if (!w_p) return Paragraph();

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_p.child( "w:pPr" );
    return Paragraph( impl );
}

Paragraph::operator bool()
{
    return impl_ != nullptr && impl_->w_p_;
}

bool dxfrg::operator==( const Paragraph& left, const Paragraph& right )
{
    if (!left.impl_ && !right.impl_) return true;
    if (left.impl_ && right.impl_) return left.impl_->w_p_ == right.impl_->w_p_;
    return false;
}

void Paragraph::operator=( const Paragraph& right )
{
    if (this == &right) return;
    if (impl_ != nullptr) delete impl_;
    if (right.impl_ != nullptr) {
        impl_ = new Impl;
        impl_->w_body_ = right.impl_->w_body_;
        impl_->w_p_ = right.impl_->w_p_;
        impl_->w_pPr_ = right.impl_->w_pPr_;
    }
    else {
        impl_ = nullptr;
    }
}

Section Paragraph::GetSection()
{
    if (!impl_) return Section();
    Section::Impl* impl = new Section::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = impl_->w_p_;
    impl->w_pPr_ = impl_->w_pPr_;
    Section section( impl );
    section.FindSectionProperties();
    return section;
}

Section Paragraph::InsertSectionBreak()
{
    Section section = GetSection();
    // this paragraph will be the last paragraph of the new section
    if (section) section.Split();
    return section;
}

Section Paragraph::RemoveSectionBreak()
{
    Section section = GetSection();
    if (section && section.IsSplit()) section.Merge();
    return section;
}

bool Paragraph::HasSectionBreak()
{
    return GetSection().IsSplit();
}