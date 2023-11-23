#include "Section.h"

#include "Paragraph.h"

#include "XMLWriter.h"

using namespace dxfrg;

std::ostream& 
dxfrg::operator<<( std::ostream& out, const Section& s )
{
    if (s.impl_) {
        xml_string_writer writer;
        s.impl_->w_p_.print( writer, "  " );
        out << writer.Result;
    }
    else {
        out << "<section />";
    }
    return out;
}

// class Section
Section::Section()
{

}

Section::Section( Impl* impl ) : impl_( impl )
{

}

Section::Section( const Section& s )
{
    impl_ = new Impl;
    impl_->w_body_ = s.impl_->w_body_;
    impl_->w_p_ = s.impl_->w_p_;
    impl_->w_p_last_ = s.impl_->w_p_last_;
    impl_->w_pPr_ = s.impl_->w_pPr_;
    impl_->w_pPr_last_ = s.impl_->w_pPr_last_;
    impl_->w_sectPr_ = s.impl_->w_sectPr_;
}

Section::~Section()
{
    if (impl_ != nullptr) {
        delete impl_;
        impl_ = nullptr;
    }
}

void Section::FindSectionProperties()
{
    if (!impl_) return;
    pugi::xml_node w_p_next = impl_->w_p_, w_p, w_pPr, w_sectPr;
    do {
        w_p = w_p_next;
        w_pPr = w_p.child( "w:pPr" );
        w_sectPr = w_pPr.child( "w:sectPr" );

        w_p_next = w_p.next_sibling( "w:p" );
    } while (w_sectPr.empty() && !w_p_next.empty());

    impl_->w_p_last_ = w_p;
    impl_->w_pPr_last_ = w_pPr;
    impl_->w_sectPr_ = w_sectPr;

    if (impl_->w_sectPr_.empty()) {
        impl_->w_sectPr_ = impl_->w_body_.child( "w:sectPr" );
    }
}

void Section::Split()
{
    if (!impl_) return;
    if (IsSplit()) return;
    impl_->w_p_last_ = impl_->w_p_;
    impl_->w_pPr_last_ = impl_->w_pPr_;
    impl_->w_sectPr_ = impl_->w_pPr_.append_copy( impl_->w_sectPr_ );
}

bool Section::IsSplit()
{
    if (!impl_) return false;
    return impl_->w_pPr_.child( "w:sectPr" );
}

void Section::Merge()
{
    if (!impl_) return;
    if (impl_->w_pPr_.child( "w:sectPr" ).empty()) return;
    impl_->w_pPr_last_.remove_child( impl_->w_sectPr_ );
    FindSectionProperties();
}

void Section::SetPageSize( const int w, const int h )
{
    if (!impl_) return;
    auto pgSz = impl_->w_sectPr_.child( "w:pgSz" );
    if (!pgSz) {
        pgSz = impl_->w_sectPr_.append_child( "w:pgSz" );
    }
    auto pgSzW = pgSz.attribute( "w:w" );
    if (!pgSzW) {
        pgSzW = pgSz.append_attribute( "w:w" );
    }
    auto pgSzH = pgSz.attribute( "w:h" );
    if (!pgSzH) {
        pgSzH = pgSz.append_attribute( "w:h" );
    }
    pgSzW.set_value( w );
    pgSzH.set_value( h );
}

void Section::GetPageSize( int& w, int& h )
{
    if (!impl_) return;
    w = h = 0;
    auto pgSz = impl_->w_sectPr_.child( "w:pgSz" );
    if (!pgSz) return;
    auto pgSzW = pgSz.attribute( "w:w" );
    if (!pgSzW) return;
    auto pgSzH = pgSz.attribute( "w:h" );
    if (!pgSzH) return;
    w = pgSzW.as_int();
    h = pgSzH.as_int();
}

void Section::SetPageOrient( const Orientation orient )
{
    if (!impl_) return;
    auto pgSz = impl_->w_sectPr_.child( "w:pgSz" );
    if (!pgSz) {
        pgSz = impl_->w_sectPr_.append_child( "w:pgSz" );
    }
    auto pgSzH = pgSz.attribute( "w:h" );
    if (!pgSzH) {
        pgSzH = pgSz.append_attribute( "w:h" );
    }
    auto pgSzW = pgSz.attribute( "w:w" );
    if (!pgSzW) {
        pgSzW = pgSz.append_attribute( "w:w" );
    }
    auto pgSzOrient = pgSz.attribute( "w:orient" );
    if (!pgSzOrient) {
        pgSzOrient = pgSz.append_attribute( "w:orient" );
    }
    int hVal = pgSzH.as_int();
    int wVal = pgSzW.as_int();
    switch (orient) {
        case Orientation::Landscape:
            if (hVal < wVal) return;
            pgSzOrient.set_value( "landscape" );
            pgSzH.set_value( wVal );
            pgSzW.set_value( hVal );
            break;
        case Orientation::Portrait:
            if (hVal > wVal) return;
            pgSzOrient.set_value( "portrait" );
            pgSzH.set_value( wVal );
            pgSzW.set_value( hVal );
            break;
    }
}

Section::Orientation Section::GetPageOrient()
{
    if (!impl_) return Orientation::Unknown;
    Orientation orient = Orientation::Portrait;
    auto pgSz = impl_->w_sectPr_.child( "w:pgSz" );
    if (!pgSz) return orient;
    auto pgSzOrient = pgSz.attribute( "w:orient" );
    if (!pgSzOrient) return orient;
    if (std::string( pgSzOrient.value() ).compare( "landscape" ) == 0) {
        orient = Orientation::Landscape;
    }
    return orient;
}

void Section::SetPageMargin( const int top, const int bottom,
                             const int left, const int right )
{
    if (!impl_) return;
    auto pgMar = impl_->w_sectPr_.child( "w:pgMar" );
    if (!pgMar) {
        pgMar = impl_->w_sectPr_.append_child( "w:pgMar" );
    }
    auto pgMarTop = pgMar.attribute( "w:top" );
    if (!pgMarTop) {
        pgMarTop = pgMar.append_attribute( "w:top" );
    }
    auto pgMarBottom = pgMar.attribute( "w:bottom" );
    if (!pgMarBottom) {
        pgMarBottom = pgMar.append_attribute( "w:bottom" );
    }
    auto pgMarLeft = pgMar.attribute( "w:left" );
    if (!pgMarLeft) {
        pgMarLeft = pgMar.append_attribute( "w:left" );
    }
    auto pgMarRight = pgMar.attribute( "w:right" );
    if (!pgMarRight) {
        pgMarRight = pgMar.append_attribute( "w:right" );
    }
    pgMarTop.set_value( top );
    pgMarBottom.set_value( bottom );
    pgMarLeft.set_value( left );
    pgMarRight.set_value( right );
}

void Section::GetPageMargin( int& top, int& bottom,
                             int& left, int& right )
{
    if (!impl_) return;
    top = bottom = left = right = 0;
    auto pgMar = impl_->w_sectPr_.child( "w:pgMar" );
    if (!pgMar) return;
    auto pgMarTop = pgMar.attribute( "w:top" );
    if (!pgMarTop) return;
    auto pgMarBottom = pgMar.attribute( "w:bottom" );
    if (!pgMarBottom) return;
    auto pgMarLeft = pgMar.attribute( "w:left" );
    if (!pgMarLeft) return;
    auto pgMarRight = pgMar.attribute( "w:right" );
    if (!pgMarRight) return;
    top = pgMarTop.as_int();
    bottom = pgMarBottom.as_int();
    left = pgMarLeft.as_int();
    right = pgMarRight.as_int();
}

void Section::SetPageMargin( const int header, const int footer )
{
    if (!impl_) return;
    auto pgMar = impl_->w_sectPr_.child( "w:pgMar" );
    if (!pgMar) {
        pgMar = impl_->w_sectPr_.append_child( "w:pgMar" );
    }
    auto pgMarHeader = pgMar.attribute( "w:header" );
    if (!pgMarHeader) {
        pgMarHeader = pgMar.append_attribute( "w:header" );
    }
    auto pgMarFooter = pgMar.attribute( "w:footer" );
    if (!pgMarFooter) {
        pgMarFooter = pgMar.append_attribute( "w:footer" );
    }
    pgMarHeader.set_value( header );
    pgMarFooter.set_value( footer );
}

void Section::GetPageMargin( int& header, int& footer )
{
    if (!impl_) return;
    header = footer = 0;
    auto pgMar = impl_->w_sectPr_.child( "w:pgMar" );
    if (!pgMar) return;
    auto pgMarHeader = pgMar.attribute( "w:header" );
    if (!pgMarHeader) return;
    auto pgMarFooter = pgMar.attribute( "w:footer" );
    if (!pgMarFooter) return;
    header = pgMarHeader.as_int();
    footer = pgMarFooter.as_int();
}

void Section::SetPageNumber( const PageNumberFormat fmt, const unsigned int start )
{
    if (!impl_) return;

    auto footerReference = impl_->w_sectPr_.child( "w:footerReference" );
    if (!footerReference) {
        footerReference = impl_->w_sectPr_.append_child( "w:footerReference" );
    }

    auto footerReferenceId = footerReference.attribute( "r:id" );
    if (!footerReferenceId) {
        footerReferenceId = footerReference.append_attribute( "r:id" );
    }
    footerReferenceId.set_value( "rId1" );

    auto footerReferenceType = footerReference.attribute( "w:type" );
    if (!footerReferenceType) {
        footerReferenceType = footerReference.append_attribute( "w:type" );
    }
    footerReferenceType.set_value( "default" );

    auto pgNumType = impl_->w_sectPr_.child( "w:pgNumType" );
    if (!pgNumType) {
        pgNumType = impl_->w_sectPr_.append_child( "w:pgNumType" );
    }

    auto pgNumTypeFmt = pgNumType.attribute( "w:fmt" );
    if (!pgNumTypeFmt) {
        pgNumTypeFmt = pgNumType.append_attribute( "w:fmt" );
    }
    const char* fmtVal = "";
    switch (fmt) {
        case PageNumberFormat::Decimal:
            fmtVal = "decimal";
            break;
        case PageNumberFormat::NumberInDash:
            fmtVal = "numberInDash";
            break;
        case PageNumberFormat::CardinalText:
            fmtVal = "cardinalText";
            break;
        case PageNumberFormat::OrdinalText:
            fmtVal = "ordinalText";
            break;
        case PageNumberFormat::LowerLetter:
            fmtVal = "lowerLetter";
            break;
        case PageNumberFormat::UpperLetter:
            fmtVal = "upperLetter";
            break;
        case PageNumberFormat::LowerRoman:
            fmtVal = "lowerRoman";
            break;
        case PageNumberFormat::UpperRoman:
            fmtVal = "upperRoman";
            break;
    }
    pgNumTypeFmt.set_value( fmtVal );

    auto pgNumTypeStart = pgNumType.attribute( "w:start" );
    if (start > 0) {
        if (!pgNumTypeStart) {
            pgNumTypeStart = pgNumType.append_attribute( "w:start" );
            pgNumTypeStart.set_value( start );
        }
    }
    else {
        if (pgNumTypeStart) {
            pgNumType.remove_attribute( pgNumTypeStart );
        }
    }
}

void Section::RemovePageNumber()
{
    if (!impl_) return;

    auto footerReference = impl_->w_sectPr_.child( "w:footerReference" );
    if (footerReference) {
        impl_->w_sectPr_.remove_child( footerReference );
    }

    auto pgNumType = impl_->w_sectPr_.child( "w:pgNumType" );
    if (pgNumType) {
        impl_->w_sectPr_.remove_child( pgNumType );
    }
}

Paragraph Section::FirstParagraph()
{
    return Prev().LastParagraph().Next();
}

Paragraph Section::LastParagraph()
{
    if (!impl_) return Paragraph();
    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = impl_->w_p_last_;
    impl->w_pPr_ = impl_->w_pPr_last_;
    return Paragraph( impl );
}

Section Section::Next()
{
    if (!impl_) return Section();
    auto w_p = impl_->w_p_last_.next_sibling();
    if (w_p.empty()) return Section();

    Section::Impl* impl = new Section::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_p.child( "w:pPr" );
    Section s( impl );
    s.FindSectionProperties();
    return s;
}

Section Section::Prev()
{
    if (!impl_) return Section();

    pugi::xml_node w_p_prev, w_p, w_pPr, w_sectPr;

    w_p_prev = impl_->w_p_.previous_sibling();
    if (w_p_prev.empty()) return Section();

    do {
        w_p = w_p_prev;
        w_pPr = w_p.child( "w:pPr" );
        w_sectPr = w_pPr.child( "w:sectPr" );
        w_p_prev = w_p.previous_sibling();
    } while (w_sectPr.empty() && !w_p_prev.empty());

    Section::Impl* impl = new Section::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = impl->w_p_last_ = w_p;
    impl->w_pPr_ = impl->w_pPr_last_ = w_pPr;
    impl->w_sectPr_ = w_sectPr;
    return Section( impl );
}

Section::operator bool()
{
    return impl_ != nullptr && impl_->w_sectPr_;
}

bool dxfrg::operator==( const Section& left, const Section& right )
{
    if (!left.impl_ && !right.impl_) return true;
    if (left.impl_ && right.impl_) return left.impl_->w_sectPr_ == right.impl_->w_sectPr_;
    return false;
}

void Section::operator=( const Section& right )
{
    if (this == &right) return;
    if (impl_ != nullptr) delete impl_;
    if (right.impl_ != nullptr) {
        impl_ = new Impl;
        impl_->w_body_ = right.impl_->w_body_;
        impl_->w_p_ = right.impl_->w_p_;
        impl_->w_p_last_ = right.impl_->w_p_last_;
        impl_->w_pPr_ = right.impl_->w_pPr_;
        impl_->w_pPr_last_ = right.impl_->w_pPr_last_;
        impl_->w_sectPr_ = right.impl_->w_sectPr_;
    }
    else {
        impl_ = nullptr;
    }
}