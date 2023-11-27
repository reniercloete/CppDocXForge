#include "Document.h"
#include "Paragraph.h"
#include "Utils.h"
#include "Section.h"
#include "Table.h"
#include "TableCell.h"
#include "Constants.h"
#include "TextFrame.h"
#include "Image.h"

void BuildBasicDocument( dxfrg::Document& Doc )
{
    auto p1 = Doc.AppendParagraph( "Hello, World!", 12, "Times New Roman" );
    //Need to use u8string?
    /*auto p2 = Doc.AppendParagraph( u8"你好，世界！", 14, u8"宋体" );
    auto p3 = Doc.AppendParagraph( u8"你好，World!", 16, "Times New Roman", u8"宋体" );*/

    auto p4 = Doc.AppendParagraph();
    p4.SetAlignment( dxfrg::Paragraph::Alignment::Centered );

    auto p4r1 = p4.AppendRun( "This is a simple sentence. ", 12, "Arial" );
    p4r1.SetCharacterSpacing( dxfrg::Pt2Twip( 2 ) );

    //Need to use u8string?
    //auto p4r2 = p4.AppendRun( u8"这是一个简单的句子。" );
    auto p4r2 = p4.AppendRun( "Blah" );
    p4r2.SetFontSize( 14 );
    p4r2.SetFontStyle( dxfrg::Run::Bold | dxfrg::Run::Italic );
}

void BuildBreaksDocument( dxfrg::Document& Doc )
{
    auto p = Doc.AppendParagraph();
    p.SetAlignment( dxfrg::Paragraph::Alignment::Left );

    auto r = p.AppendRun();
    r.AppendText( "This is" );
    r.AppendLineBreak();
    r.AppendText( "a simple sentence." );
    p.AppendPageBreak();

    Doc.AppendParagraph( "see you next page." );
}

void BuildPageNumberDocument( dxfrg::Document& Doc )
{
    auto p1 = Doc.AppendParagraph( "This is the 3rd paragraph in the 1st section." );
    auto p2 = Doc.AppendParagraph( "This is the 4th paragraph in the 1st section." );
    p2.InsertSectionBreak();

    auto p3 = Doc.AppendParagraph( "This is the 5th paragraph in the 2nd section." );
    auto p4 = Doc.AppendParagraph( "This is the 6th paragraph in the 2nd section." );

    auto s1 = p1.GetSection();
    auto s2 = p3.GetSection();

    s1.SetPageNumber( dxfrg::Section::PageNumberFormat::NumberInDash );
    s2.SetPageNumber( dxfrg::Section::PageNumberFormat::UpperLetter, 1 );
}

void BuildParagraphDocument( dxfrg::Document& Doc )
{
    Doc.AppendParagraph( "This is the 2nd paragraph." );
    Doc.AppendParagraph( "This is the 3rd paragraph." );
    Doc.AppendParagraph( "This is the 4th paragraph." );
    Doc.AppendParagraph( "This is the 5th paragraph." );
    Doc.PrependParagraph( "This is the 1st paragraph." );

    auto p1 = Doc.FirstParagraph();
    auto p2 = p1.Next();
    auto p3 = p2.Next();
    auto p5 = Doc.LastParagraph();
    auto p4 = p5.Prev();

    std::cout << p1.GetText() << std::endl;
    std::cout << p2.GetText() << std::endl;
    std::cout << p3.GetText() << std::endl;
    std::cout << p4.GetText() << std::endl;
    std::cout << p5.GetText() << std::endl;

    Doc.RemoveParagraph( p2 );
    Doc.InsertParagraphBefore( p4 ).AppendRun( "New paragraph before the 4th paragraph." );
    Doc.InsertParagraphAfter( p4 ).AppendRun( "New paragraph after the 4th paragraph." );
}

void BuildRunDocument( dxfrg::Document& Doc )
{
    auto p = Doc.AppendParagraph();
    //Need to use u8string?
    //auto r = p.AppendRun( u8"你好，World!", 16, "Times New Roman", "Microsoft YaHei UI" );
    auto r = p.AppendRun( "Hello World!", 16, "Times New Roman", "Microsoft YaHei UI" );

    r.SetCharacterSpacing( dxfrg::Pt2Twip( 2 ) );

    r.SetFontStyle( dxfrg::Run::Bold | dxfrg::Run::Underline );
    r.SetFontStyle( dxfrg::Run::Bold | dxfrg::Run::Italic );
    auto fontStyle = r.GetFontStyle();
    r.SetFontStyle( fontStyle | dxfrg::Run::Strikethrough );

    auto fontSize = r.GetFontSize();
    std::cout << "Font Size: " << fontSize << std::endl;

    auto characterSpacing = dxfrg::Twip2Pt( r.GetCharacterSpacing() );
    std::cout << "Character Spacing: " << characterSpacing << std::endl;

    std::string fontAscii, fontEastAsia;
    r.GetFont( fontAscii, fontEastAsia );
    std::cout << "Font Ascii: " << fontAscii << std::endl;
    std::cout << "Font East Asia: " << fontEastAsia << std::endl;
}

void BuildSectionDocument( dxfrg::Document& Doc )
{
    auto p1 = Doc.AppendParagraph( "This is the 1st paragraph of the 1st section (A3)." );
    auto p2 = Doc.AppendSectionBreak();
    p2.AppendRun( "This is the 2nd paragraph of the 1st section (A3)." );

    auto p3 = Doc.AppendParagraph( "This is the 3rd paragraph of the 2nd section." );
    auto p4 = Doc.AppendParagraph( "This is the 4th paragraph of the 2nd section." );

    auto p5 = Doc.AppendParagraph( "This is the 5th paragraph of the 3rd section (Landscape)." );
    auto p6 = Doc.AppendParagraph( "This is the 6th paragraph of the 3rd section (Landscape)." );
    p6.InsertSectionBreak();

    auto p7 = Doc.AppendParagraph( "This is the 7th paragraph of the 4th section." );
    auto p8 = Doc.AppendParagraph( "This is the 8th paragraph of the 4th section." );

    auto s1 = p1.GetSection();
    auto s2 = p4.InsertSectionBreak();
    auto s3 = s2.Next();
    auto s4 = Doc.LastSection();

    std::cout << s1.LastParagraph().GetText() << std::endl;
    std::cout << s2.LastParagraph().GetText() << std::endl;
    std::cout << s3.LastParagraph().GetText() << std::endl;
    std::cout << s4.LastParagraph().GetText() << std::endl;

    auto firstSection = Doc.FirstSection();
    if (s1 == firstSection) {
        std::cout << "They're the same Section\n";
    }

    for (auto s = Doc.LastSection(); s; s = s.Prev()) {
        std::cout << s.LastParagraph().GetText() << std::endl;
    }

    s1.SetPageSize( dxfrg::Inch2Twip( dxfrg::A3_W ), dxfrg::Inch2Twip( dxfrg::A3_H ) ); // A3
    s3.SetPageOrient( dxfrg::Section::Orientation::Landscape );

    int w, h;
    s2.GetPageSize( w, h );
    std::cout << "Page Size: "
        << dxfrg::Twip2MM( w ) << "mm"
        << " * "
        << dxfrg::Twip2MM( h ) << "mm\n";

    auto orient = s2.GetPageOrient() == dxfrg::Section::Orientation::Landscape
        ? "Landscape"
        : "Portrait";
    std::cout << "Orientation: " << orient << std::endl;

    s4.SetPageMargin( dxfrg::CM2Twip( 2.54 ), dxfrg::CM2Twip( 2.54 ),
                      dxfrg::CM2Twip( 3.175 ), dxfrg::CM2Twip( 3.175 ) );
    s4.SetPageMargin( dxfrg::CM2Twip( 1.5 ), dxfrg::CM2Twip( 1.75 ) );

    std::cout << s2.LastParagraph().GetText() << std::endl;

    p2.SetAlignment( dxfrg::Paragraph::Alignment::Left );
    p4.SetAlignment( dxfrg::Paragraph::Alignment::Centered );
    p6.SetAlignment( dxfrg::Paragraph::Alignment::Right );
}

void BuildSpacingIndentDocument( dxfrg::Document& Doc )
{
#define text1 "I feel that there is much to be said for the Celtic belief that the souls of those whom we have lost are held captive in some inferior being, in an animal, in a plant, in some inanimate object, and so effectively lost to us until the day (which to many never comes) when we happen to pass by the tree or to obtain possession of the object which forms their prison."
#define text2 "Then they start and tremble, they call us by our name, and as soon as we have recognized their voice the spell is broken. We have delivered them: they have overcome death and return to share our life. And so it is with our own past. It is a labor in vain to attempt to recapture it: all the efforts of our intellect must prove futile."
#define text3 "The past is hidden somewhere outside the realm, beyond the reach of intellect, in some material object (in the sensation which that material object will give us) which we do not suspect. And as for that object, it depends on chance whether we come upon it or not before we ourselves must die."
    // Page 1
    auto p1 = Doc.AppendParagraph( text1 );
    auto p2 = Doc.AppendParagraph( text2 );
    auto p3 = Doc.AppendParagraph( text3 );
    auto p4 = Doc.AppendParagraph( text1 );
    auto p5 = Doc.AppendParagraph( text2 );
    auto p6 = Doc.AppendParagraph( text3 );

    p2.SetFontStyle( dxfrg::Run::Bold );
    p4.SetFontStyle( dxfrg::Run::Bold );
    p6.SetFontStyle( dxfrg::Run::Bold );

    // Line spacing
    p1.SetLineSpacingSingle();                   // Single
    p2.SetLineSpacingLines( 1.5 );                 // 1.5 lines
    p3.SetLineSpacingLines( 2 );                   // Double (2 lines)
    p4.SetLineSpacingAtLeast( dxfrg::Pt2Twip( 12 ) ); // At Least (12 pt)
    p5.SetLineSpacingExactly( dxfrg::Pt2Twip( 12 ) ); // Exactly (12 pt)
    p6.SetLineSpacingLines( 3 );                   // Multiple (3 lines)

    // Indent
    p1.SetLeftIndentChars( 2 );
    p2.SetRightIndent( dxfrg::CM2Twip( 3 ) );
    p3.SetFirstLineChars( 2 );
    p4.SetFirstLine( dxfrg::CM2Twip( 2 ) );
    p5.SetHangingChars( 2 );
    p6.SetHanging( dxfrg::CM2Twip( 2 ) );

    // Page 2
    Doc.AppendPageBreak();
    auto p7 = Doc.AppendParagraph( "This is the 7th paragraph." );
    auto p8 = Doc.AppendParagraph( "This is the 8th paragraph." );
    auto p9 = Doc.AppendParagraph( "This is the 9th paragraph." );
    auto p10 = Doc.AppendParagraph( "This is the 10th paragraph." );

    // Spacing
    p8.SetBeforeSpacingAuto();
    p9.SetBeforeSpacingLines( 1.5 );
    p9.SetAfterSpacing( dxfrg::Pt2Twip( 10 ) );
}

void BuildTableDocument( dxfrg::Document& Doc )
{
    auto tbl = Doc.AppendTable( 4, 5 );

    tbl.GetCell( 0, 0 ).FirstParagraph().AppendRun( "AAA" );
    tbl.GetCell( 0, 1 ).FirstParagraph().AppendRun( "BBB" );
    tbl.GetCell( 0, 2 ).FirstParagraph().AppendRun( "CCC" );
    tbl.GetCell( 0, 3 ).FirstParagraph().AppendRun( "DDD" );

    tbl.GetCell( 1, 0 ).FirstParagraph().AppendRun( "EEE" );
    tbl.GetCell( 1, 1 ).FirstParagraph().AppendRun( "FFF" );

    tbl.SetAlignment( dxfrg::Table::Alignment::Centered );

    tbl.SetTopBorders( dxfrg::Table::BorderStyle::Single, 1, "FF0000" );
    tbl.SetBottomBorders( dxfrg::Table::BorderStyle::Dotted, 2, "00FF00" );
    tbl.SetLeftBorders( dxfrg::Table::BorderStyle::Dashed, 3, "0000FF" );
    tbl.SetRightBorders( dxfrg::Table::BorderStyle::DotDash, 1, "FFFF00" );
    tbl.SetInsideHBorders( dxfrg::Table::BorderStyle::Double, 1, "FF00FF" );
    tbl.SetInsideVBorders( dxfrg::Table::BorderStyle::Wave, 1, "00FFFF" );

    tbl.SetCellMarginLeft( dxfrg::CM2Twip( 0.19 ) );
    tbl.SetCellMarginRight( dxfrg::CM2Twip( 0.19 ) );
}

void BuildAdvancedTableDocument( dxfrg::Document& Doc )
{
    Doc.AppendParagraph( "Table 1" );
    auto t1 = Doc.AppendTable( 4, 5 );

    auto c11 = t1.GetCell( 1, 1 );
    auto c12 = t1.GetCell( 1, 2 );
    if (t1.MergeCells( c11, c12 )) {
        std::cout << "c11 c12 merged\n";
    }

    auto c04 = t1.GetCell( 0, 4 );
    auto c14 = t1.GetCell( 1, 4 );
    if (t1.MergeCells( c04, c14 )) {
        std::cout << "c04 c14 merged\n";
    }

    auto c11_12 = t1.GetCell( 1, 2 );
    auto c13 = t1.GetCell( 1, 3 );
    if (t1.MergeCells( c12, c13 )) {
        std::cout << "c11_12 c13 merged\n";
    }

    auto c10 = t1.GetCell( 1, 0 );
    auto c11_12_13 = t1.GetCell( 1, 1 );
    if (t1.MergeCells( c10, c13 )) {
        std::cout << "c10 c11_12_13 merged\n";
    }

    auto c24 = t1.GetCell( 2, 4 );
    auto c34 = t1.GetCell( 3, 4 );
    if (t1.MergeCells( c24, c34 )) {
        std::cout << "c24 c34 merged\n";
    }

    if (t1.MergeCells( c24, c04 )) {
        std::cout << "c24 c04 merged\n";
    }

    t1.GetCell( 0, 0 ).FirstParagraph().AppendRun( "AAA" );
    t1.GetCell( 0, 1 ).FirstParagraph().AppendRun( "BBB" );
    t1.GetCell( 0, 2 ).FirstParagraph().AppendRun( "CCC" );
    t1.GetCell( 0, 3 ).FirstParagraph().AppendRun( "DDD" );
    t1.GetCell( 1, 0 ).FirstParagraph().AppendRun( "FFF" );

    c04.SetVerticalAlignment( dxfrg::TableCell::Alignment::Center );
    auto c04p1 = c04.FirstParagraph();
    c04p1.SetAlignment( dxfrg::Paragraph::Alignment::Centered );
    c04p1.AppendRun( "EEE" );

    Doc.AppendParagraph( "Table 2" );
    auto t2 = Doc.AppendTable( 4, 4 );

    t2.MergeCells( t2.GetCell( 0, 1 ), t2.GetCell( 1, 1 ) );
    t2.MergeCells( t2.GetCell( 0, 1 ), t2.GetCell( 2, 1 ) );
    t2.MergeCells( t2.GetCell( 0, 2 ), t2.GetCell( 1, 2 ) );
    t2.MergeCells( t2.GetCell( 0, 2 ), t2.GetCell( 2, 2 ) );
    t2.MergeCells( t2.GetCell( 0, 1 ), t2.GetCell( 2, 2 ) );

    t2.MergeCells( t2.GetCell( 0, 0 ), t2.GetCell( 1, 0 ) );
    t2.MergeCells( t2.GetCell( 1, 0 ), t2.GetCell( 2, 0 ) );

    t2.MergeCells( t2.GetCell( 0, 0 ), t2.GetCell( 0, 1 ) );
}

void BuildTextFrameDocument( dxfrg::Document& Doc )
{
    Doc.AppendParagraph( "Hello, World!" );

    auto frame = Doc.AppendTextFrame( dxfrg::CM2Twip( 4 ), dxfrg::CM2Twip( 5 ) );
    frame.AppendRun( "This is the text frame paragraph." );
    frame.SetPositionX( dxfrg::TextFrame::Position::Left, dxfrg::TextFrame::Anchor::Page );
    frame.SetPositionY( dxfrg::TextFrame::Position::Top, dxfrg::TextFrame::Anchor::Margin );
    // frame.SetPositionX(CM2Twip(1), TextFrame::Anchor::Margin);
    // frame.SetPositionY(CM2Twip(1), TextFrame::Anchor::Margin);
    frame.SetTextWrapping( dxfrg::TextFrame::Wrapping::Around );
}

void BuildTraverseDocument( dxfrg::Document& Doc )
{
    // Section 1
    Doc.AppendParagraph( "This is the 1st paragraph." );
    Doc.AppendParagraph( "This is the 2nd paragraph." );
    Doc.AppendParagraph( "This is the 3rd paragraph." );

    auto p = Doc.AppendSectionBreak();
    p.AppendRun( "This is the 1st run. " );
    p.AppendRun( "This is the 2nd run. " );

    auto r = p.AppendRun();
    r.AppendText( "This is the 1st text. " );
    r.AppendText( "This is the 2nd text. " );
    r.AppendText( "This is the 3rd text." );

    // Section 2
    Doc.AppendParagraph( "This is the 5th paragraph." );
    Doc.AppendParagraph( "This is the 6th paragraph." );
    Doc.AppendSectionBreak().AppendRun( "This is the 7th paragraph." );

    // Section 3
    Doc.AppendParagraph( "This is the 8th paragraph." );
    Doc.AppendParagraph( "This is the 9th paragraph." );


    for (auto p = Doc.FirstParagraph(); p; p = p.Next()) {
        std::cout << "paragraph: \n";
        for (auto r = p.FirstRun(); r; r = r.Next()) {
            std::cout << "run: " << r.GetText() << std::endl;
        }
    }

    auto s2 = Doc.FirstSection().Next();
    auto last = s2.LastParagraph();
    for (auto p = s2.FirstParagraph(); p; p = p.Next()) {
        p.SetAlignment( dxfrg::Paragraph::Alignment::Centered );
        if (p == last) break;
    }
}

void BuildImageDocument( dxfrg::Document& Doc )
{
    Doc.AppendImage( 7.960431, 6.239756944, "../bin/Example.png" );
    Doc.AppendImage( 15.92086111, 12.47951389, "../bin/Example.emf" );
}

void BuildSuperSubDocument( dxfrg::Document& Doc )
{
    // Section 1
    auto p = Doc.AppendParagraph();

    auto r = p.AppendRun();
    r.AppendText( "This is the 1" );
    r = p.AppendRun();
    r.SetVerticalAlign( dxfrg::Run::SubScript );
    r.AppendText( "st" );
    r = p.AppendRun();
    r.AppendText( " line containing a subscript." );
    r.AppendLineBreak();

    r = p.AppendRun();
    r.AppendText( "This is the 2" );
    r = p.AppendRun();
    r.SetVerticalAlign( dxfrg::Run::SuperScript );
    r.AppendText( "nd" );
    r = p.AppendRun();
    r.AppendText( " line containing a superscript." );
                  
}

int main()
{
    dxfrg::Document Doc( "../bin/Test.docx" );

    //BuildBasicDocument( Doc );
    //BuildBreaksDocument( Doc );
    //BuildPageNumberDocument( Doc );
    //BuildParagraphDocument( Doc );
    //BuildRunDocument( Doc );
    //BuildSectionDocument( Doc );
    //BuildSpacingIndentDocument( Doc );
    //BuildTableDocument( Doc );
    //BuildAdvancedTableDocument( Doc );
    //BuildTextFrameDocument( Doc );
    //BuildTraverseDocument( Doc );
    //BuildImageDocument( Doc );
    BuildSuperSubDocument( Doc );
    Doc.Save();
}
