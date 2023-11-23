#include "Document.h"

//#include "DocumentBoilerPlatePar.h"
#include "XMLWriter.h"
#include "Paragraph.h"
#include "Section.h"
#include "Table.h"
#include "TextFrame.h"
#include "Image.h"
#include "Utils.h"

#include "zip.h"

#include <filesystem>

using namespace dxfrg;

std::ostream&
dxfrg::operator<<( std::ostream& out, const Document& doc )
{
    if (doc.impl_)
    {
        xml_string_writer writer;
        doc.impl_->w_body_.print( writer, "  " );
        out << writer.Result;
    }
    else
    {
        out << "<document />";
    }
    return out;
}

//bool
//CDocument::Save()
//{
//    /*if (!impl_) return false;
//
//    xml_string_writer writer;
//    impl_->doc_.save( writer, "", pugi::format_raw );
//    const char* buf = writer.result.c_str();*/
//
//    struct zip_t* zip = zip_open( mFilePath.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'w');
//    if (zip == nullptr) {
//        return false;
//    }
//
//    zip_entry_open( zip, "_rels/.rels" );
//    zip_entry_write( zip, _RELS, std::strlen( _RELS ) );
//    zip_entry_close( zip );
//
//    /*zip_entry_open(zip, "word/document.xml");
//    zip_entry_write( zip, buf, std::strlen( buf ) );
//    zip_entry_close( zip );*/
//
//    zip_entry_open( zip, "word/document.xml" );
//    zip_entry_write( zip, DOCUMENT_XML, std::strlen( DOCUMENT_XML ) );
//    zip_entry_close( zip );
//
//    zip_entry_open( zip, "word/_rels/document.xml.rels" );
//    zip_entry_write( zip, DOCUMENT_XML_RELS, std::strlen( DOCUMENT_XML_RELS ) );
//    zip_entry_close( zip );
//
//    /*zip_entry_open( zip, "word/fontTable.xml" );
//    zip_entry_write( zip, DOCUMENT_XML_FONT_TABLE, std::strlen( DOCUMENT_XML_FONT_TABLE ) );
//    zip_entry_close( zip );*/
//
//    zip_entry_open( zip, "docProps/app.xml" );
//    zip_entry_write( zip, DOCPROPS_APP, std::strlen( DOCPROPS_APP ) );
//    zip_entry_close( zip );
//
//    zip_entry_open( zip, "docProps/core.xml" );
//    zip_entry_write( zip, DOCPROPS_CORE, std::strlen( DOCPROPS_CORE ) );
//    zip_entry_close( zip );
//
//    zip_entry_open( zip, "[Content_Types].xml" );
//    zip_entry_write( zip, CONTENT_TYPES_XML, std::strlen( CONTENT_TYPES_XML ) );
//    zip_entry_close( zip );
//
//    zip_close( zip );
//    return true;
//}

#define _RELS R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>)"
#define DOCUMENT_XML R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:oel="http://schemas.microsoft.com/office/2019/extlst" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14"><w:body><w:sectPr><w:pgSz w:w="11906" w:h="16838" /><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0" /><w:cols w:space="425" /><w:docGrid w:type="lines" w:linePitch="312" /></w:sectPr></w:body></w:document>)"
#define CONTENT_TYPES_XML R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="emf" ContentType="image/x-emf"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" /><Default Extension="xml" ContentType="application/xml" /><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" /><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml" /></Types>)"
#define DOCUMENT_XML_RELS R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml" /></Relationships>)"
#define FOOTER1_XML R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:oel="http://schemas.microsoft.com/office/2019/extlst" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14"><w:p><w:pPr><w:jc w:val="center" /></w:pPr><w:r><w:fldChar w:fldCharType="begin" /></w:r><w:r><w:instrText>PAGE \* MERGEFORMAT</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end" /></w:r></w:p></w:ftr>)"

// class Document
Document::Document( const std::string& path )
{
    impl_ = new Impl;
    impl_->doc_.load_buffer( DOCUMENT_XML, std::strlen( DOCUMENT_XML ), pugi::parse_declaration );
    impl_->w_body_ = impl_->doc_.child( "w:document" ).child( "w:body" );
    impl_->w_sectPr_ = impl_->w_body_.child( "w:sectPr" );
    impl_->path_ = path;
}

Document::~Document()
{
    if (impl_ != nullptr) {
        delete impl_;
        impl_ = nullptr;
    }
}

bool Document::Save()
{
    if (!impl_) return false;

    xml_string_writer doc_writer;
    impl_->doc_.save( doc_writer, "", pugi::format_raw );
    const char* buf = doc_writer.Result.c_str();

    struct zip_t* zip = zip_open( impl_->path_.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'w' );
    if (zip == nullptr) 
    {
        return false;
    }

    zip_entry_open( zip, "_rels/.rels" );
    zip_entry_write( zip, _RELS, std::strlen( _RELS ) );
    zip_entry_close( zip );

    zip_entry_open( zip, "word/document.xml" );
    zip_entry_write( zip, buf, std::strlen( buf ) );
    zip_entry_close( zip );

    pugi::xml_document doc_rels;
    doc_rels.load_string( DOCUMENT_XML_RELS );

    auto rels = doc_rels.child( "Relationships" );
    std::string media_path = "media/";
    for (const auto& Iter : mIDToImagePathMap)
    {
        auto rel = rels.append_child( "Relationship" );
        rel.append_attribute( "Id" ) = Iter.first.c_str();
        rel.append_attribute( "Type" ) = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        std::filesystem::path p( Iter.second );
        
        rel.append_attribute( "Target" ) = (media_path + p.filename().string()).c_str();

        zip_entry_open( zip, (std::string("word/media/")+ p.filename().string()).c_str() );
        zip_entry_fwrite( zip, Iter.second.c_str() );
        zip_entry_close( zip );
        
    }

    xml_string_writer doc_rels_writer;
    doc_rels.save( doc_rels_writer, "", pugi::format_raw );
    buf = doc_rels_writer.Result.c_str();

    zip_entry_open( zip, "word/_rels/document.xml.rels" );
    zip_entry_write( zip, buf, std::strlen( buf ) ); 
    //zip_entry_write( zip, DOCUMENT_XML_RELS, std::strlen( DOCUMENT_XML_RELS ) );
    zip_entry_close( zip );

    zip_entry_open( zip, "word/footer1.xml" );
    zip_entry_write( zip, FOOTER1_XML, std::strlen( FOOTER1_XML ) );
    zip_entry_close( zip );

    zip_entry_open( zip, "[Content_Types].xml" );
    zip_entry_write( zip, CONTENT_TYPES_XML, std::strlen( CONTENT_TYPES_XML ) );
    zip_entry_close( zip );

    zip_close( zip );
    return true;
}

bool Document::Open( const std::string& path )
{
    if (!impl_) return false;

    struct zip_t* zip = zip_open( path.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'r' );
    if (zip == nullptr) {
        return false;
    }

    if (zip_entry_open( zip, "word/document.xml" ) < 0) {
        zip_close( zip );
        return false;
    }
    void* buf = nullptr;
    size_t bufsize;
    zip_entry_read( zip, &buf, &bufsize );
    zip_entry_close( zip );
    zip_close( zip );

    impl_->doc_.load_buffer( buf, bufsize, pugi::parse_declaration );
    impl_->w_body_ = impl_->doc_.child( "w:document" ).child( "w:body" );
    impl_->w_sectPr_ = impl_->w_body_.child( "w:sectPr" );
    std::free( buf );
    return true;
}

Paragraph Document::FirstParagraph()
{
    if (!impl_) return Paragraph();
    auto w_p = impl_->w_body_.child( "w:p" );
    if (!w_p) return Paragraph();

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_p.child( "w:pPr" );
    return Paragraph( impl );
}

Paragraph Document::LastParagraph()
{
    if (!impl_) return Paragraph();
    auto w_p = GetLastChild( impl_->w_body_, "w:p" );
    if (!w_p) return Paragraph();

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_p.child( "w:pPr" );
    return Paragraph( impl );
}

Section Document::FirstSection()
{
    Paragraph firstParagraph = FirstParagraph();
    if (firstParagraph) return firstParagraph.GetSection();

    Section::Impl* impl = new Section::Impl;
    impl->w_body_ = impl_->w_body_;
    Section section( impl );
    section.FindSectionProperties();
    return section;
}

Section Document::LastSection()
{
    Paragraph lastParagraph = LastParagraph();
    if (lastParagraph) return lastParagraph.GetSection();

    Section::Impl* impl = new Section::Impl;
    impl->w_body_ = impl_->w_body_;
    Section section( impl );
    section.FindSectionProperties();
    return section;
}

Paragraph Document::AppendParagraph()
{
    if (!impl_) return Paragraph();

    auto w_p = impl_->w_body_.insert_child_before( "w:p", impl_->w_sectPr_ );
    auto w_pPr = w_p.append_child( "w:pPr" );

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_pPr;
    return Paragraph( impl );
}

Paragraph Document::AppendParagraph( const std::string& text )
{
    auto p = AppendParagraph();
    p.AppendRun( text );
    return p;
}

Paragraph Document::AppendParagraph( const std::string& text,
                                     const double fontSize )
{
    auto p = AppendParagraph();
    p.AppendRun( text, fontSize );
    return p;
}

Paragraph Document::AppendParagraph( const std::string& text,
                                     const double fontSize,
                                     const std::string& fontAscii,
                                     const std::string& fontEastAsia )
{
    auto p = AppendParagraph();
    p.AppendRun( text, fontSize, fontAscii, fontEastAsia );
    return p;
}

Paragraph Document::PrependParagraph()
{
    if (!impl_) return Paragraph();

    auto w_p = impl_->w_body_.prepend_child( "w:p" );
    auto w_pPr = w_p.append_child( "w:pPr" );

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_pPr;
    return Paragraph( impl );
}

Paragraph Document::PrependParagraph( const std::string& text )
{
    auto p = PrependParagraph();
    p.AppendRun( text );
    return p;
}

Paragraph Document::PrependParagraph( const std::string& text,
                                      const double fontSize )
{
    auto p = PrependParagraph();
    p.AppendRun( text, fontSize );
    return p;
}

Paragraph Document::PrependParagraph( const std::string& text,
                                      const double fontSize,
                                      const std::string& fontAscii,
                                      const std::string& fontEastAsia )
{
    auto p = PrependParagraph();
    p.AppendRun( text, fontSize, fontAscii, fontEastAsia );
    return p;
}

Paragraph Document::InsertParagraphBefore( const Paragraph& p )
{
    if (!impl_) return Paragraph();

    auto w_p = impl_->w_body_.insert_child_before( "w:p", p.impl_->w_p_ );
    auto w_pPr = w_p.append_child( "w:pPr" );

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_pPr;
    return Paragraph( impl );
}

Paragraph Document::InsertParagraphAfter( const Paragraph& p )
{
    if (!impl_) return Paragraph();

    auto w_p = impl_->w_body_.insert_child_after( "w:p", p.impl_->w_p_ );
    auto w_pPr = w_p.append_child( "w:pPr" );

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_pPr;
    return Paragraph( impl );
}

bool Document::RemoveParagraph( Paragraph& p )
{
    if (!impl_) return false;
    return impl_->w_body_.remove_child( p.impl_->w_p_ );
}

Paragraph Document::AppendPageBreak()
{
    auto p = AppendParagraph();
    p.AppendPageBreak();
    return p;
}

Paragraph Document::AppendSectionBreak()
{
    auto p = AppendParagraph();
    p.InsertSectionBreak();
    return p;
}

Table Document::AppendTable( const int rows, const int cols )
{
    if (!impl_) return Table();

    auto w_tbl = impl_->w_body_.insert_child_before( "w:tbl", impl_->w_sectPr_ );
    auto w_tblPr = w_tbl.append_child( "w:tblPr" );
    auto w_tblGrid = w_tbl.append_child( "w:tblGrid" );

    Table::Impl* impl = new Table::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_tbl_ = w_tbl;
    impl->w_tblPr_ = w_tblPr;
    impl->w_tblGrid_ = w_tblGrid;
    Table tbl( impl );

    tbl.Create_( rows, cols );
    tbl.SetAllBorders();
    tbl.SetWidthPercent( 100 );
    tbl.SetCellMarginLeft( CM2Twip( 0.19 ) );
    tbl.SetCellMarginRight( CM2Twip( 0.19 ) );
    tbl.SetAlignment( Table::Alignment::Centered );
    return tbl;
}

void Document::RemoveTable( Table& tbl )
{
    if (!impl_) return;
    impl_->w_body_.remove_child( tbl.impl_->w_tbl_ );
}

Image 
Document::AppendImage( /*const*/ double w, /*const*/ double h,
                       const std::string& path )
{
    if (!impl_) return Image();
    ++mIDCounter;

    auto w_p = impl_->w_body_.insert_child_before( "w:p", impl_->w_sectPr_ );
    auto w_pPr = w_p.append_child( "w:pPr" );
    auto w_r = w_p.append_child( "w:r" );

    Paragraph::Impl* paragraph_impl = new Paragraph::Impl;
    paragraph_impl->w_body_ = impl_->w_body_;
    paragraph_impl->w_p_ = w_p;
    paragraph_impl->w_pPr_ = w_pPr;

    auto w_drawing = w_r.append_child( "w:drawing" );
    auto w_inline = w_drawing.append_child( "wp:inline" );
	w_inline.append_attribute( "distT" ) = "0";
	w_inline.append_attribute( "distB" ) = "0";
	w_inline.append_attribute( "distL" ) = "0";
	w_inline.append_attribute( "distR" ) = "0";
   
    auto w_extent = w_inline.append_child( "wp:extent" );
    w_extent.append_attribute( "cx" ) = std::to_string( CM2Emu( w ) ).c_str();
    w_extent.append_attribute( "cy" ) = std::to_string( CM2Emu( h ) ).c_str();
    
    /*auto w_effectExtent = w_inline.append_child("wp:effectExtent");
	w_effectExtent.append_attribute( "l" ) = "0";
	w_effectExtent.append_attribute( "t" ) = "0";
	w_effectExtent.append_attribute( "r" ) = "2540";
    w_effectExtent.append_attribute( "b" ) = "3175";*/

	auto w_docPr = w_inline.append_child( "wp:docPr" );
    w_docPr.append_attribute( "id" ) = std::to_string( mIDCounter ).c_str();
	w_docPr.append_attribute( "name" ) = "Picture 1";
	w_docPr.append_attribute( "descr" ) = "Picture 1";

	/*auto wp_cNvGraphicFramePr = w_inline.append_child( "cNvGraphicFramePr" );
	auto a_graphicFrameLocks = wp_cNvGraphicFramePr.append_child( "a:graphicFrameLocks" );
	a_graphicFrameLocks.append_attribute( "xmlns:a" ) = "http://schemas.openxmlformats.org/drawingml/2006/main";
	a_graphicFrameLocks.append_attribute( "noChangeAspect" ) = "1";*/

	auto wp_graphic = w_inline.append_child( "a:graphic" );
    wp_graphic.append_attribute( "xmlns:a" ) = "http://schemas.openxmlformats.org/drawingml/2006/main";
	auto a_graphicData = wp_graphic.append_child( "a:graphicData" );
	a_graphicData.append_attribute( "uri" ) = "http://schemas.openxmlformats.org/drawingml/2006/picture";

	auto pic_pic = a_graphicData.append_child( "pic:pic" );
	pic_pic.append_attribute( "xmlns:pic" ) = "http://schemas.openxmlformats.org/drawingml/2006/picture";

	auto pic_nvPicPr = pic_pic.append_child( "pic:nvPicPr" );
	auto pic_cNvPr = pic_nvPicPr.append_child( "pic:cNvPr" );
	pic_cNvPr.append_attribute( "id" ) = std::to_string( mIDCounter ).c_str();
	pic_cNvPr.append_attribute( "name" ) = "Picture 1";
	pic_cNvPr.append_attribute( "descr" ) = "Picture 1";
    auto pic_cNvPicPr = pic_nvPicPr.append_child( "pic:cNvPicPr" );
    pic_cNvPicPr.text().set( "" );

    std::string TextID = "rId" + std::to_string( mIDCounter + 1 );;
    auto pic_blipFill = pic_pic.append_child( "pic:blipFill" );
    auto a_blip = pic_blipFill.append_child( "a:blip" );
    a_blip.append_attribute( "r:embed" ) = TextID.c_str();
    auto a_stretch = pic_blipFill.append_child( "a:stretch" );
    auto a_fillRect = a_stretch.append_child( "a:fillRect" );

    auto pic_spPr = pic_pic.append_child( "pic:spPr" );
    auto a_xfrm = pic_spPr.append_child( "a:xfrm" );
    auto a_off = a_xfrm.append_child( "a:off" );
    a_off.append_attribute( "x" ) = "0";
    a_off.append_attribute( "y" ) = "0";
    auto a_ext = a_xfrm.append_child( "a:ext" );
    a_ext.append_attribute( "cx" ) = std::to_string( CM2Emu( w ) ).c_str();
    a_ext.append_attribute( "cy" ) = std::to_string( CM2Emu( h ) ).c_str();

    auto a_prstGeom = pic_spPr.append_child( "a:prstGeom" );
    a_prstGeom.append_attribute( "prst" ) = "rect";
    auto a_avLst = a_prstGeom.append_child( "a:avLst" );/**/

    Image::Impl* impl = new Image::Impl;
    //impl->w_framePr_ = w_framePr;
    auto image = Image( impl, paragraph_impl );

    mIDToImagePathMap[TextID] = path;

    /*textFrame.SetSize( w, h );
    textFrame.SetBorders();*/
    return image;
}

TextFrame Document::AppendTextFrame( const int w, const int h )
{
    if (!impl_) return TextFrame();

    auto w_p = impl_->w_body_.insert_child_before( "w:p", impl_->w_sectPr_ );
    auto w_pPr = w_p.append_child( "w:pPr" );
    auto w_framePr = w_pPr.append_child( "w:framePr" );


    Paragraph::Impl* paragraph_impl = new Paragraph::Impl;
    paragraph_impl->w_body_ = impl_->w_body_;
    paragraph_impl->w_p_ = w_p;
    paragraph_impl->w_pPr_ = w_pPr;

    TextFrame::Impl* impl = new TextFrame::Impl;
    impl->w_framePr_ = w_framePr;
    auto textFrame = TextFrame( impl, paragraph_impl );

    textFrame.SetSize( w, h );
    textFrame.SetBorders();
    return textFrame;
}
