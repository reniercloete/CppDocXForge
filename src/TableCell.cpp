#include "TableCell.h"

#include "Paragraph.h"

using namespace dxfrg;

// class TableCell
TableCell::TableCell()
{

}

TableCell::TableCell( Impl* impl ) : impl_( impl )
{

}

TableCell::TableCell( const TableCell& tc )
{
    impl_ = new Impl;
    impl_->c_ = tc.impl_->c_;
    impl_->w_tr_ = tc.impl_->w_tr_;
    impl_->w_tc_ = tc.impl_->w_tc_;
    impl_->w_tcPr_ = tc.impl_->w_tcPr_;
    impl_->w_trPr_ = tc.impl_->w_trPr_;
}

TableCell::~TableCell()
{
    if (impl_ != nullptr) {
        delete impl_;
        impl_ = nullptr;
    }
}

void TableCell::operator=( const TableCell& right )
{
    if (this == &right) return;
    if (impl_ != nullptr) delete impl_;
    if (right.impl_ != nullptr) {
        impl_ = new Impl;
        impl_->c_ = right.impl_->c_;
        impl_->w_tr_ = right.impl_->w_tr_;
        impl_->w_tc_ = right.impl_->w_tc_;
        impl_->w_tcPr_ = right.impl_->w_tcPr_;
        impl_->w_trPr_ = right.impl_->w_trPr_;
    }
    else {
        impl_ = nullptr;
    }
}

TableCell::operator bool()
{
    return impl_ != nullptr && impl_->w_tc_;
}

bool TableCell::empty() const
{
    return impl_ == nullptr || !impl_->w_tc_;
}

void TableCell::SetHeight( const int h, const char* units )
{
    if (!impl_) return;
    auto w_trHeight = impl_->w_trPr_.child( "w:trHeight" );
    if (!w_trHeight) {
        w_trHeight = impl_->w_trPr_.append_child( "w:trHeight" );
    }

    auto w_val = w_trHeight.attribute( "w:val" );
    if (!w_val) {
        w_val = w_trHeight.append_attribute( "w:val" );
    }

    auto w_hRule = w_trHeight.attribute( "w:hRule" );
    if (!w_hRule) {
        w_hRule = w_trHeight.append_attribute( "w:hRule" );
    }

    w_val.set_value( h );
    w_hRule.set_value( units );
}

void TableCell::SetWidth( const int w, const char* units )
{
    if (!impl_) return;
    auto w_tcW = impl_->w_tcPr_.child( "w:tcW" );
    if (!w_tcW) {
        w_tcW = impl_->w_tcPr_.append_child( "w:tcW" );
    }

    auto w_w = w_tcW.attribute( "w:w" );
    if (!w_w) {
        w_w = w_tcW.append_attribute( "w:w" );
    }

    auto w_type = w_tcW.attribute( "w:type" );
    if (!w_type) {
        w_type = w_tcW.append_attribute( "w:type" );
    }

    w_w.set_value( w );
    w_type.set_value( units );
}

void TableCell::SetVerticalAlignment( const Alignment align )
{
    if (!impl_) return;
    auto w_vAlign = impl_->w_tcPr_.child( "w:vAlign" );
    if (!w_vAlign) {
        w_vAlign = impl_->w_tcPr_.append_child( "w:vAlign" );
    }

    auto w_val = w_vAlign.attribute( "w:val" );
    if (!w_val) {
        w_val = w_vAlign.append_attribute( "w:val" );
    }

    const char* val = "";
    switch (align) {
        case Alignment::Top:
            val = "top";
            break;
        case Alignment::Center:
            val = "center";
            break;
        case Alignment::Bottom:
            val = "bottom";
            break;
    }
    w_val.set_value( val );
}

void TableCell::SetCellSpanning_( const int cols )
{
    if (!impl_) return;
    auto w_gridSpan = impl_->w_tcPr_.child( "w:gridSpan" );
    if (cols == 1) {
        if (w_gridSpan) {
            impl_->w_tcPr_.remove_child( w_gridSpan );
        }
        return;
    }
    if (!w_gridSpan) {
        w_gridSpan = impl_->w_tcPr_.append_child( "w:gridSpan" );
    }

    auto w_val = w_gridSpan.attribute( "w:val" );
    if (!w_val) {
        w_val = w_gridSpan.append_attribute( "w:val" );
    }

    w_val.set_value( cols );
}

Paragraph TableCell::AppendParagraph()
{
    if (!impl_) return Paragraph();
    auto w_p = impl_->w_tc_.append_child( "w:p" );
    auto w_pPr = w_p.append_child( "w:pPr" );

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_tc_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_pPr;
    return Paragraph( impl );
}

Paragraph TableCell::FirstParagraph()
{
    if (!impl_) return Paragraph();
    auto w_p = impl_->w_tc_.child( "w:p" );
    auto w_pPr = w_p.child( "w:pPr" );

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_tc_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_pPr;
    return Paragraph( impl );
}
