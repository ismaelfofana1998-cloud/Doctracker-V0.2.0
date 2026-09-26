using System;
using System.Drawing;
using System.Windows.Forms;
using Doctracker.Core.Services;

namespace Doctracker.AddIn.UI
{
    internal sealed partial class DocumentCanvas
    {
        private TableGrid tableGrid;
        private RectangleF tableZone;
        private int tablePage,tableHandle=-1;
        private bool tableColumn;
        public event Action TableGridChanged;
        // null: add on top/left borders only; true/false: add a column/row anywhere inside.
        public bool? TableAddColumn {get;set;}
        public void EditTableGrid(int page,RectangleF zone,TableGrid grid)
        {
            tableGrid=grid??throw new ArgumentNullException(nameof(grid));tablePage=page;tableZone=zone;tableHandle=-1;
            ClearSelection();GoToPage(page);normalizedSelection=zone;RevealSelection();InvalidatePages();
        }
        private RectangleF TableBounds(Size size)=>new RectangleF(tableZone.X*size.Width,tableZone.Y*size.Height,tableZone.Width*size.Width,tableZone.Height*size.Height);
        private int TableHandleAt(Point point,RectangleF box,out bool column)
        {
            var radius=Math.Max(7,Font.Height/2f);
            for(var i=1;i<tableGrid.Columns.Count-1;i++)
                if(Math.Abs(point.X-(box.Left+(float)tableGrid.Columns[i]*box.Width))<=radius &&
                    (Math.Abs(point.Y-box.Top)<=radius || Math.Abs(point.Y-box.Bottom)<=radius)){column=true;return i;}
            for(var i=1;i<tableGrid.Rows.Count-1;i++)
                if(Math.Abs(point.Y-(box.Top+(float)tableGrid.Rows[i]*box.Height))<=radius &&
                    (Math.Abs(point.X-box.Left)<=radius || Math.Abs(point.X-box.Right)<=radius)){column=false;return i;}
            column=false;return -1;
        }
        private bool TableMouseDown(MouseEventArgs e)
        {
            if(tableGrid==null)return false;
            if(CurrentPageNumber!=tablePage)return true;
            var box=TableBounds(picture.Size);var index=TableHandleAt(e.Location,box,out var column);
            if(e.Button==MouseButtons.Right)
            {
                if(index>0 && tableGrid.Remove(column,index))TableChanged();
                return true;
            }
            if(e.Button!=MouseButtons.Left)return true;
            if(index>0){tableHandle=index;tableColumn=column;picture.Capture=true;return true;}
            var near=box;near.Inflate(8,8);if(!near.Contains(e.Location))return true;
            bool? add=Math.Abs(e.Y-box.Top)<=8?true:Math.Abs(e.X-box.Left)<=8?false:TableAddColumn;
            if(add.HasValue)
            {
                var at=add.Value?(e.X-box.X)/box.Width:(e.Y-box.Y)/box.Height;
                try {if(tableGrid.Add(add.Value,at))TableChanged();}
                catch(InvalidOperationException error){MessageBox.Show(this,error.Message,"Tableau");}
            }
            return true;
        }
        private bool TableMouseMove(MouseEventArgs e)
        {
            if(tableGrid==null)return false;
            if(tableHandle>0 && CurrentPageNumber==tablePage)
            {
                var box=TableBounds(picture.Size);
                var at=tableColumn?(e.X-box.X)/box.Width:(e.Y-box.Y)/box.Height;
                if(tableGrid.Move(tableColumn,tableHandle,at))picture.Invalidate();
            }
            else if(CurrentPageNumber==tablePage)
            {
                var hit=TableHandleAt(e.Location,TableBounds(picture.Size),out var column);
                picture.Cursor=hit>0?(column?Cursors.SizeWE:Cursors.SizeNS):Cursors.Cross;
            }
            return true;
        }
        private bool TableMouseUp(MouseEventArgs e)
        {
            if(tableGrid==null)return false;
            if(tableHandle>0){tableHandle=-1;picture.Capture=false;TableChanged();}
            return true;
        }
        private void TableChanged(){InvalidatePages();TableGridChanged?.Invoke();}
        private bool PaintTableGrid(Graphics graphics,Size size,int page)
        {
            if(tableGrid==null)return false;
            if(page!=tablePage)return true;
            var box=TableBounds(size);var radius=Math.Max(4,Font.Height/3f);
            using(var pen=new Pen(SnipTheme.ColorFor(Doctracker.Core.Models.SnipType.Table),1.5f))
            using(var brush=new SolidBrush(pen.Color))
            {
                graphics.DrawRectangle(pen,box.X,box.Y,box.Width,box.Height);
                for(var i=1;i<tableGrid.Columns.Count-1;i++)
                {
                    var x=box.Left+(float)tableGrid.Columns[i]*box.Width;graphics.DrawLine(pen,x,box.Top,x,box.Bottom);
                    graphics.FillEllipse(brush,x-radius,box.Top-radius,2*radius,2*radius);graphics.FillEllipse(brush,x-radius,box.Bottom-radius,2*radius,2*radius);
                }
                for(var i=1;i<tableGrid.Rows.Count-1;i++)
                {
                    var y=box.Top+(float)tableGrid.Rows[i]*box.Height;graphics.DrawLine(pen,box.Left,y,box.Right,y);
                    graphics.FillEllipse(brush,box.Left-radius,y-radius,2*radius,2*radius);graphics.FillEllipse(brush,box.Right-radius,y-radius,2*radius,2*radius);
                }
            }
            return true;
        }
    }
}
