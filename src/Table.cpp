#include "Table.h"

#include "TableCell.h"
#include "Utils.h"
#include "Paragraph.h"

using namespace dxfrg;

// class Table
Table::Table( Impl* impl ) : impl_( impl )
{

}

Table::Table()
{

}

Table::Table( const Table& t )
{
    impl_ = new Impl;
    impl_->w_body_ = t.impl_->w_body_;
    impl_->w_tbl_ = t.impl_->w_tbl_;
    impl_->w_tblPr_ = t.impl_->w_tblPr_;
    impl_->w_tblGrid_ = t.impl_->w_tblGrid_;
    impl_->rows_ = t.impl_->rows_;
    impl_->cols_ = t.impl_->cols_;
    impl_->grid_ = t.impl_->grid_;
}

Table::~Table()
{
    if (impl_ != nullptr) {
        delete impl_;
        impl_ = nullptr;
    }
}

void Table::operator=( const Table& right )
{
    if (this == &right) return;
    if (impl_ != nullptr) delete impl_;
    if (right.impl_ != nullptr) {
        impl_ = new Impl;
        impl_->w_body_ = right.impl_->w_body_;
        impl_->w_tbl_ = right.impl_->w_tbl_;
        impl_->w_tblPr_ = right.impl_->w_tblPr_;
        impl_->w_tblGrid_ = right.impl_->w_tblGrid_;
        impl_->rows_ = right.impl_->rows_;
        impl_->cols_ = right.impl_->cols_;
        impl_->grid_ = right.impl_->grid_;
    }
    else {
        impl_ = nullptr;
    }
}

void Table::Create_( const int rows, const int cols )
{
    if (!impl_) return;
    impl_->rows_ = rows;
    impl_->cols_ = cols;

    // init grid
    for (int i = 0; i < rows; i++) {
        Row row;
        for (int j = 0; j < cols; j++) {
            Cell cell = { i, j, 1, 1 };
            row.push_back( cell );
        }
        impl_->grid_.push_back( row );
    }

    // init table
    for (int i = 0; i < rows; i++) {
        auto w_gridCol = impl_->w_tblGrid_.append_child( "w:gridCol" );
        auto w_tr = impl_->w_tbl_.append_child( "w:tr" );
        auto w_trPr = w_tr.append_child( "w:trPr" );

        for (int j = 0; j < cols; j++) {
            auto w_tc = w_tr.append_child( "w:tc" );
            auto w_tcPr = w_tc.append_child( "w:tcPr" );

            TableCell::Impl* impl = new TableCell::Impl;
            impl->c_ = &impl_->grid_[i][j];
            impl->w_tr_ = w_tr;
            impl->w_trPr_ = w_trPr;
            impl->w_tc_ = w_tc;
            impl->w_tcPr_ = w_tcPr;
            TableCell tc( impl );
            // A table cell must contain at least one block-level element, 
            // even if it is an empty <p/>.
            tc.AppendParagraph();
        }
    }
}

TableCell Table::GetCell( const int row, const int col )
{
    if (!impl_) return TableCell();
    if (row < 0 || row >= impl_->rows_ || col < 0 || col >= impl_->cols_) {
        return TableCell();
    }

    Cell* c = &impl_->grid_[row][col];
    return GetCell_( c->row, c->col );
}

TableCell Table::GetCell_( const int row, const int col )
{
    if (!impl_) return TableCell();
    int i = 0;
    auto w_tr = impl_->w_tbl_.child( "w:tr" );
    while (i < row && !w_tr.empty()) {
        i++;
        w_tr = w_tr.next_sibling( "w:tr" );
    }
    if (w_tr.empty()) {
        return TableCell();
    }

    int j = 0;
    auto w_tc = w_tr.child( "w:tc" );
    while (j < col && !w_tc.empty()) {
        j += impl_->grid_[row][j].cols;
        w_tc = w_tc.next_sibling( "w:tc" );
    }
    if (w_tc.empty()) {
        return TableCell();
    }

    TableCell::Impl* impl = new TableCell::Impl;
    impl->c_ = &impl_->grid_[row][col];
    impl->w_tr_ = w_tr;
    impl->w_tc_ = w_tc;
    impl->w_tcPr_ = w_tc.child( "w:tcPr" );
    impl->w_trPr_ = w_tr.child( "w:trPr" );
    return TableCell( impl );
}

bool Table::MergeCells( TableCell tc1, TableCell tc2 )
{
    if (tc1.empty() || tc2.empty()) {
        return false;
    }

    Cell* c1 = tc1.impl_->c_;
    Cell* c2 = tc2.impl_->c_;

    if (c1->row == c2->row && c1->col != c2->col && c1->rows == c2->rows) {
        Cell* left_cell, * right_cell;
        TableCell* left_tc, * right_tc;
        if (c1->col < c2->col) {
            left_cell = c1;
            left_tc = &tc1;
            right_cell = c2;
            right_tc = &tc2;
        }
        else {
            left_cell = c2;
            left_tc = &tc2;
            right_cell = c1;
            right_tc = &tc1;
        }

        const int col = left_cell->col;
        if ((right_cell->col - col) == left_cell->cols) {
            const int cols = left_cell->cols + right_cell->cols;

            // update right grid
            const int right_col = right_cell->col;
            const int right_cols = right_cell->cols;
            for (int i = 0; i < right_cell->rows; i++) {
                const int y = right_cell->row + i;
                for (int j = 0; j < right_cols; j++) {
                    Cell& c = impl_->grid_[y][right_col + j];
                    c.col = col;
                    c.cols = cols;
                }
            }

            // update cells
            for (int i = 0; i < right_cell->rows; i++) {
                RemoveCell_( GetCell_( right_cell->row + i, right_cell->col ) );
            }
            for (int i = 0; i < left_cell->rows; i++) {
                GetCell_( left_cell->row + i, left_cell->col ).SetCellSpanning_( cols );
            }

            // update left grid
            const int left_cols = left_cell->cols;
            for (int i = 0; i < left_cell->rows; i++) {
                const int y = left_cell->row + i;
                for (int j = 0; j < left_cols; j++) {
                    Cell& c = impl_->grid_[y][left_cell->col + j];
                    c.cols = cols;
                }
            }

            right_tc->impl_->c_ = left_cell;
            right_tc->impl_->w_tc_ = left_tc->impl_->w_tc_;
            right_tc->impl_->w_tcPr_ = left_tc->impl_->w_tcPr_;
            return true;
        }
    }
    else if (c1->col == c2->col && c1->row != c2->row && c1->cols == c2->cols) {
        Cell* top_cell, * bottom_cell;
        TableCell* top_tc, * bottom_tc;
        if (c1->row < c2->row) {
            top_cell = c1;
            top_tc = &tc1;
            bottom_cell = c2;
            bottom_tc = &tc2;
        }
        else {
            top_cell = c2;
            top_tc = &tc2;
            bottom_cell = c1;
            bottom_tc = &tc1;
        }

        const int row = top_cell->row;
        if ((bottom_cell->row - top_cell->row) == top_cell->rows) {
            const int rows = top_cell->rows + bottom_cell->rows;

            // update cells
            if (top_cell->rows == 1) {
                auto w_vMerge = top_tc->impl_->w_tcPr_.append_child( "w:vMerge" );
                auto w_val = w_vMerge.append_attribute( "w:val" );
                w_val.set_value( "restart" );
            }
            if (bottom_cell->rows == 1) {
                bottom_tc->impl_->w_tcPr_.append_child( "w:vMerge" );
            }
            else {
                bottom_tc->impl_->w_tcPr_.remove_child( "w:vMerge" );
                bottom_tc->impl_->w_tcPr_.append_child( "w:vMerge" );
            }

            // update top grid
            const int top_rows = top_cell->rows;
            for (int i = 0; i < top_rows; i++) {
                const int x = top_cell->row + i;
                for (int j = 0; j < top_cell->cols; j++) {
                    Cell& c = impl_->grid_[x][top_cell->col + j];
                    c.rows = rows;
                }
            }

            // update bottom grid
            const int bottom_row = bottom_cell->row;
            const int bottom_rows = bottom_cell->rows;
            for (int i = 0; i < bottom_rows; i++) {
                const int x = bottom_row + i;
                for (int j = 0; j < bottom_cell->cols; j++) {
                    Cell& c = impl_->grid_[x][bottom_cell->col + j];
                    c.row = row;
                    c.rows = rows;
                }
            }

            bottom_tc->impl_->c_ = top_cell;
            bottom_tc->impl_->w_tc_ = top_tc->impl_->w_tc_;
            bottom_tc->impl_->w_tcPr_ = top_tc->impl_->w_tcPr_;
            bottom_tc->impl_->w_trPr_ = top_tc->impl_->w_trPr_;
            return true;
        }
    }

    return false;
}

void Table::RemoveCell_( TableCell tc )
{
    if (!impl_) return;
    tc.impl_->w_tr_.remove_child( tc.impl_->w_tc_ );
}

void Table::SetWidthAuto()
{
    SetWidth( 0, "auto" );
}

void Table::SetWidthPercent( const double w )
{
    SetWidth( w / 0.02, "pct" );
}

void Table::SetWidth( const int w, const char* units )
{
    if (!impl_) return;
    auto w_tblW = impl_->w_tblPr_.child( "w:tblW" );
    if (!w_tblW) {
        w_tblW = impl_->w_tblPr_.append_child( "w:tblW" );
    }

    auto w_w = w_tblW.attribute( "w:w" );
    if (!w_w) {
        w_w = w_tblW.append_attribute( "w:w" );
    }

    auto w_type = w_tblW.attribute( "w:type" );
    if (!w_type) {
        w_type = w_tblW.append_attribute( "w:type" );
    }

    w_w.set_value( w );
    w_type.set_value( units );
}

void Table::SetCellMarginTop( const int w, const char* units )
{
    SetCellMargin( "w:top", w, units );
}

void Table::SetCellMarginBottom( const int w, const char* units )
{
    SetCellMargin( "w:bottom", w, units );
}

void Table::SetCellMarginLeft( const int w, const char* units )
{
    SetCellMargin( "w:start", w, units );
}

void Table::SetCellMarginRight( const int w, const char* units )
{
    SetCellMargin( "w:end", w, units );
}

void Table::SetCellMargin( const char* elemName, const int w, const char* units )
{
    if (!impl_) return;
    auto w_tblCellMar = impl_->w_tblPr_.child( "w:tblCellMar" );
    if (!w_tblCellMar) {
        w_tblCellMar = impl_->w_tblPr_.append_child( "w:tblCellMar" );
    }

    auto w_tblCellMarChild = w_tblCellMar.child( elemName );
    if (!w_tblCellMarChild) {
        w_tblCellMarChild = w_tblCellMar.append_child( elemName );
    }

    auto w_w = w_tblCellMarChild.attribute( "w:w" );
    if (!w_w) {
        w_w = w_tblCellMarChild.append_attribute( "w:w" );
    }

    auto w_type = w_tblCellMarChild.attribute( "w:type" );
    if (!w_type) {
        w_type = w_tblCellMarChild.append_attribute( "w:type" );
    }

    w_w.set_value( w );
    w_type.set_value( units );
}

void Table::SetAlignment( const Alignment alignment )
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
    }

    auto w_jc = impl_->w_tblPr_.child( "w:jc" );
    if (!w_jc) {
        w_jc = impl_->w_tblPr_.append_child( "w:jc" );
    }
    auto w_val = w_jc.attribute( "w:val" );
    if (!w_val) {
        w_val = w_jc.append_attribute( "w:val" );
    }
    w_val.set_value( val );
}

void Table::SetTopBorders( const BorderStyle style, const double width, const char* color )
{
    SetBorders_( "w:top", style, width, color );
}

void Table::SetBottomBorders( const BorderStyle style, const double width, const char* color )
{
    SetBorders_( "w:bottom", style, width, color );
}

void Table::SetLeftBorders( const BorderStyle style, const double width, const char* color )
{
    SetBorders_( "w:start", style, width, color );
}

void Table::SetRightBorders( const BorderStyle style, const double width, const char* color )
{
    SetBorders_( "w:end", style, width, color );
}

void Table::SetInsideHBorders( const BorderStyle style, const double width, const char* color )
{
    SetBorders_( "w:insideH", style, width, color );
}

void Table::SetInsideVBorders( const BorderStyle style, const double width, const char* color )
{
    SetBorders_( "w:insideV", style, width, color );
}

void Table::SetInsideBorders( const BorderStyle style, const double width, const char* color )
{
    SetInsideHBorders( style, width, color );
    SetInsideVBorders( style, width, color );
}

void Table::SetOutsideBorders( const BorderStyle style, const double width, const char* color )
{
    SetTopBorders( style, width, color );
    SetBottomBorders( style, width, color );
    SetLeftBorders( style, width, color );
    SetRightBorders( style, width, color );
}

void Table::SetAllBorders( const BorderStyle style, const double width, const char* color )
{
    SetOutsideBorders( style, width, color );
    SetInsideBorders( style, width, color );
}

void Table::SetBorders_( const char* elemName, const BorderStyle style, const double width, const char* color )
{
    if (!impl_) return;
    auto w_tblBorders = impl_->w_tblPr_.child( "w:tblBorders" );
    if (!w_tblBorders) {
        w_tblBorders = impl_->w_tblPr_.append_child( "w:tblBorders" );
    }
    SetBorders( w_tblBorders, elemName, style, width, color );
}