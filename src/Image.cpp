#include "Image.h"

using namespace dxfrg;

// class Image
Image::Image()
{

}

Image::Image( Impl* impl, Paragraph::Impl* paragraph_impl )
    : Paragraph( paragraph_impl ), impl_( impl )
{

}

Image::Image( const Image& tf )
    : Paragraph( tf )
{
    impl_ = new Impl;
    //impl_->w_framePr_ = tf.impl_->w_framePr_;
}

Image::~Image()
{
    if (impl_ != nullptr) {
        delete impl_;
        impl_ = nullptr;
    }
}

//void Image::SetSize( const int w, const int h )
//{
//    if (!impl_) return;
//    auto w_w = impl_->w_framePr_.attribute( "w:w" );
//    if (!w_w) {
//        w_w = impl_->w_framePr_.append_attribute( "w:w" );
//    }
//    auto w_h = impl_->w_framePr_.attribute( "w:h" );
//    if (!w_h) {
//        w_h = impl_->w_framePr_.append_attribute( "w:h" );
//    }
//
//    w_w.set_value( w );
//    w_h.set_value( h );
//}
//
//void Image::SetPositionX( const Position align, const Anchor ralativeTo )
//{
//    SetAnchor_( "w:hAnchor", ralativeTo );
//    SetPosition_( "w:xAlign", align );
//}
//
//void Image::SetPositionY( const Position align, const Anchor ralativeTo )
//{
//    SetAnchor_( "w:vAnchor", ralativeTo );
//    SetPosition_( "w:yAlign", align );
//}
//
//void Image::SetPositionX( const int x, const Anchor ralativeTo )
//{
//    SetAnchor_( "w:hAnchor", ralativeTo );
//    SetPosition_( "w:x", x );
//}
//
//void Image::SetPositionY( const int y, const Anchor ralativeTo )
//{
//    SetAnchor_( "w:vAnchor", ralativeTo );
//    SetPosition_( "w:y", y );
//}
//
//void Image::SetAnchor_( const char* attrName, const Anchor anchor )
//{
//    if (!impl_) return;
//    auto w_anchor = impl_->w_framePr_.attribute( attrName );
//    if (!w_anchor) {
//        w_anchor = impl_->w_framePr_.append_attribute( attrName );
//    }
//
//    const char* val = "";
//    switch (anchor) {
//        case Anchor::Page:
//            val = "page";
//            break;
//        case Anchor::Margin:
//            val = "margin";
//            break;
//    }
//    w_anchor.set_value( val );
//}
//
//void Image::SetPosition_( const char* attrName, const Position align )
//{
//    if (!impl_) return;
//    auto w_align = impl_->w_framePr_.attribute( attrName );
//    if (!w_align) {
//        w_align = impl_->w_framePr_.append_attribute( attrName );
//    }
//
//    const char* val = "";
//    switch (align) {
//        case Position::Left:
//            val = "left";
//            break;
//        case Position::Center:
//            val = "center";
//            break;
//        case Position::Right:
//            val = "right";
//            break;
//        case Position::Top:
//            val = "top";
//            break;
//        case Position::Bottom:
//            val = "bottom";
//            break;
//    }
//    w_align.set_value( val );
//}
//
//void Image::SetPosition_( const char* attrName, const int twip )
//{
//    if (!impl_) return;
//    auto w_Align = impl_->w_framePr_.attribute( attrName );
//    if (!w_Align) {
//        w_Align = impl_->w_framePr_.append_attribute( attrName );
//    }
//    w_Align.set_value( twip );
//}
//
//void Image::SetTextWrapping( const Wrapping wrapping )
//{
//    if (!impl_) return;
//    auto w_wrap = impl_->w_framePr_.attribute( "w:wrap" );
//    if (!w_wrap) {
//        w_wrap = impl_->w_framePr_.append_attribute( "w:wrap" );
//    }
//
//    const char* val = "";
//    switch (wrapping) {
//        case Wrapping::Around:
//            val = "around";
//            break;
//        case Wrapping::None:
//            val = "none";
//            break;
//    }
//    w_wrap.set_value( val );
//}