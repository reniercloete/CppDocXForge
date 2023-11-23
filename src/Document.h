#pragma once

#include "pugixml.hpp"

#include <iostream>
#include <map>
#include <string>

namespace dxfrg
{
	/*class Document
	{
		friend std::ostream& operator<<( std::ostream& out, const Document& doc );
	public:
		Document( const std::string& FilePath ) :
			mFilePath( FilePath )
		{}

		bool Save();
	private:
		std::string mFilePath = "";
	};*/

    class Paragraph;
    class Section;
    class Table;
    class TextFrame;
    class Image;

    class Document
    {
        friend std::ostream& operator<<( std::ostream& out, const Document& doc );

    public:
        // constructs an empty document
        Document( const std::string& path );
        ~Document();

        // save document to file
        bool Save();
        bool Open( const std::string& path );

        // get paragraph
        Paragraph FirstParagraph();
        Paragraph LastParagraph();

        // add paragraph
        Paragraph AppendParagraph();
        Paragraph AppendParagraph( const std::string& text );
        Paragraph AppendParagraph( const std::string& text, const double fontSize );
        Paragraph AppendParagraph( const std::string& text, const double fontSize, const std::string& fontAscii, const std::string& fontEastAsia = "" );
        Paragraph PrependParagraph();
        Paragraph PrependParagraph( const std::string& text );
        Paragraph PrependParagraph( const std::string& text, const double fontSize );
        Paragraph PrependParagraph( const std::string& text, const double fontSize, const std::string& fontAscii, const std::string& fontEastAsia = "" );

        Paragraph InsertParagraphBefore( const Paragraph& p );
        Paragraph InsertParagraphAfter( const Paragraph& p );
        bool RemoveParagraph( Paragraph& p );

        Paragraph AppendPageBreak();

        // get section
        Section FirstSection();
        Section LastSection();

        // add section
        Paragraph AppendSectionBreak();

        // add table
        Table AppendTable( const int rows, const int cols );
        void RemoveTable( Table& tbl );

        // add text frame
        TextFrame AppendTextFrame( const int w, const int h );

        // add text frame
        Image AppendImage( /*const*/ double w, /*const*/ double h,
                           const std::string& path );

    private:
        struct Impl;
        Impl* impl_ = nullptr;

        int mIDCounter = 0;
        std::map<std::string, std::string> mIDToImagePathMap;
    }; // class Document

    struct Document::Impl
    {
        std::string        path_;
        pugi::xml_document doc_;
        pugi::xml_node     w_body_;
        pugi::xml_node     w_sectPr_;
    };

}