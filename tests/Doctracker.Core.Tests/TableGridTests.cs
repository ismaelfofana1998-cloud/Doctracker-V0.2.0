using System;
using System.Linq;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class TableGridTests
    {
        private static WordRecord Word(string text,double x,double y,int line=1)=>new WordRecord {Text=text,X=x,Y=y,Width=.1,Height=.05,Line=line};
        [Fact] public void Initial_grid_keeps_columns_rows_and_blank_cells()
        {
            var words=new[]{Word("Article",.1,.1),Word("Montant",.65,.1),Word("00123",.1,.5,2)};
            var grid=TableGrid.FromWords(words);var cells=grid.Extract(words);
            Assert.Equal(2,grid.RowCount);Assert.Equal(2,grid.ColumnCount);
            Assert.Equal(new[]{"Article","Montant","00123",""},cells.Select(c=>c.Text));
            Assert.Equal(0,cells[0].X);Assert.Equal(1,cells[3].X+cells[3].Width,8);
        }
        [Fact] public void Moving_a_separator_reassigns_each_word_once_without_rerunning_ocr()
        {
            var words=new[]{Word("LEFT",.1,.1),Word("RIGHT",.7,.1)};
            var grid=new TableGrid();grid.Add(true,.5);
            Assert.Equal(new[]{"LEFT","RIGHT"},grid.Extract(words).Select(c=>c.Text));
            grid.Move(true,1,.85);
            Assert.Equal(new[]{"LEFT RIGHT",""},grid.Extract(words).Select(c=>c.Text));
            Assert.Equal(.7,words[1].X);
        }
        [Fact] public void Removed_row_separator_merges_lines_and_retains_source_geometry()
        {
            var words=new[]{Word("TOP",.1,.1),Word("BOTTOM",.1,.7,2)};
            var grid=new TableGrid();grid.Add(false,.5);Assert.Equal(2,grid.Extract(words).Count);
            grid.Remove(false,1);var cell=Assert.Single(grid.Extract(words));
            Assert.Equal("TOP\nBOTTOM",cell.Text);Assert.Equal(1,cell.Height);
        }
        [Fact] public void A_word_on_the_boundary_is_assigned_once_and_multiline_order_is_preserved()
        {
            var grid=new TableGrid();grid.Add(true,.5);
            var words=new[]{Word("B",.7,.4,2),Word("A",.45,.1)};
            var cells=grid.Extract(words);Assert.Equal("",cells[0].Text);Assert.Equal("A\nB",cells[1].Text);
        }
        [Fact] public void Borders_cannot_be_deleted_or_crossed_and_duplicate_points_are_ignored()
        {
            var grid=new TableGrid();Assert.True(grid.Add(true,.4));Assert.True(grid.Add(true,.6));
            Assert.False(grid.Add(true,.4));Assert.False(grid.Add(true,0));Assert.False(grid.Add(false,double.NaN));
            Assert.False(grid.Remove(true,0));Assert.False(grid.Move(true,3,.8));
            grid.Move(true,1,.9);Assert.True(grid.Columns[1]<grid.Columns[2]);
            Assert.Equal(0,grid.Columns[0]);Assert.Equal(1,grid.Columns.Last());
        }
        [Fact] public void Grid_cell_limit_prevents_unbounded_excel_output()
        {
            var grid=new TableGrid();
            for(var i=1;i<100;i++){grid.Add(true,i/100d);grid.Add(false,i/100d);}
            Assert.Equal(10000,grid.RowCount*grid.ColumnCount);
            Assert.Throws<InvalidOperationException>(()=>grid.Add(true,.005));
            Assert.Equal(10000,grid.Extract(new WordRecord[0]).Count);
        }
    }
}
