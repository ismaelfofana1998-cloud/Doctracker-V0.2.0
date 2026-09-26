using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Doctracker.Core.Models;
using Doctracker.Core.Services;

public static class TableEditorProbe
{
    private const BindingFlags Flags=BindingFlags.Instance|BindingFlags.NonPublic|BindingFlags.Public;
    private static void Check(bool value,string message){if(!value)throw new Exception(message);}
    private static object Field(object instance,string name){return instance.GetType().GetField(name,Flags).GetValue(instance);}
    private static object Call(object instance,string name,params object[] args){return instance.GetType().GetMethod(name,Flags).Invoke(instance,args);}
    private static void Mouse(object canvas,string name,MouseButtons button,double x,double y,RectangleF zone)
    {
        var picture=(PictureBox)Field(canvas,"picture");
        var point=new Point((int)Math.Round((zone.X+x*zone.Width)*picture.Width),(int)Math.Round((zone.Y+y*zone.Height)*picture.Height));
        Call(canvas,name,picture,new MouseEventArgs(button,1,point.X,point.Y,0));
    }
    public static string Run(Assembly assembly,string path,PageTextRecord native,string previews)
    {
        var pane=assembly.GetType("Doctracker.AddIn.UI.DoctrackerPaneControl",true);
        var extract=pane.GetMethod("ExtractPageSelection",BindingFlags.Static|BindingFlags.NonPublic);
        var zone=new RectangleF(.05f,.05f,.9f,.8f);
        var recognized=(PageTextRecord)extract.Invoke(null,new object[]{native,zone});
        Check(recognized!=null && recognized.Words.Count>0,"Native selection has no table words");
        var dialogType=assembly.GetType("Doctracker.AddIn.UI.TablePreviewDialog",true);
        using(var dialog=(Form)dialogType.GetConstructors(Flags)[0].Invoke(new object[]{path,1,zone,recognized.Words}))
        {
            dialog.Show();Application.DoEvents();
            var canvas=Field(dialog,"canvas");var preview=(DataGridView)Field(dialog,"grid");
            var layout=(TableGrid)dialogType.GetProperty("GridDefinition").GetValue(dialog,null);
            var oldColumns=layout.ColumnCount;var oldRows=layout.RowCount;var snips=0;
            EventHandler selected=(s,e)=>snips++;
            canvas.GetType().GetEvent("SelectionCompleted").AddEventHandler(canvas,selected);
            Mouse(canvas,"Picture_MouseDown",MouseButtons.Left,.94,0,zone);
            Mouse(canvas,"Picture_MouseUp",MouseButtons.Left,.94,0,zone);
            Check(layout.ColumnCount==oldColumns+1 && preview.ColumnCount==layout.ColumnCount,"Clicking the top edge did not add a column and update preview");
            var index=layout.Columns.Count-2;var desired=(layout.Columns[index-1]+layout.Columns[index])/2;
            Mouse(canvas,"Picture_MouseDown",MouseButtons.Left,layout.Columns[index],0,zone);
            Mouse(canvas,"Picture_MouseMove",MouseButtons.Left,desired,0,zone);
            Mouse(canvas,"Picture_MouseUp",MouseButtons.Left,desired,0,zone);
            Check(Math.Abs(layout.Columns[index]-desired)<.01,"Dragging a column point did not move its separator");
            Mouse(canvas,"Picture_MouseDown",MouseButtons.Left,0,.94,zone);
            Mouse(canvas,"Picture_MouseUp",MouseButtons.Left,0,.94,zone);
            Check(layout.RowCount==oldRows+1 && preview.RowCount==layout.RowCount,"Clicking the left edge did not add a row");
            Mouse(canvas,"Picture_MouseDown",MouseButtons.Right,0,layout.Rows[layout.Rows.Count-2],zone);
            Mouse(canvas,"Picture_MouseUp",MouseButtons.Right,0,.94,zone);
            Check(layout.RowCount==oldRows,"Right clicking a row point did not remove it");
            preview.Rows[0].Cells[0].Value="CORRIGE";Call(dialog,"RebuildPreview");
            Check(Convert.ToString(preview.Rows[0].Cells[0].Value)=="CORRIGE","Unchanged grid discarded corrected text");
            var before=layout.Columns.ToArray();Call(canvas,"SetZoom",.55,false);Call(canvas,"GoToPage",1);Application.DoEvents();
            Check(before.SequenceEqual(layout.Columns),"Zoom changed table coordinates");
            Check(snips==0,"Editing table points created ordinary snips");
            foreach(DataGridViewColumn column in preview.Columns)Check(column.SortMode==DataGridViewColumnSortMode.NotSortable,"Sorting can detach table evidence from source rows");
            Check(preview.Height>=100 && ((Control)canvas).Height>=180,"Table editor panes collapsed");
            Call(canvas,"FitWidth");Application.DoEvents();
            using(var bitmap=new Bitmap(dialog.Width,dialog.Height))
            {dialog.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(previews,"table-editor-"+(IntPtr.Size==4?"x86":"x64")+".png"));}
            dialog.DialogResult=DialogResult.Cancel;dialog.Close();
        }
        var theme=assembly.GetType("Doctracker.AddIn.UI.SnipTheme",true);
        Check((string)theme.GetMethod("LabelFor",BindingFlags.Static|BindingFlags.Public).Invoke(null,new object[]{SnipType.Exception})=="Anomalie","French anomaly label missing");
        foreach(var kind in new[]{SnipType.Validation,SnipType.Exception})
        {
            var state=new ProjectState();var doc=new DocumentRecord {PageCount=1};state.Documents.Add(doc);
            var service=new SnipService(new ProjectStore(Path.Combine(Path.GetTempPath(),Guid.NewGuid().ToString("N"))),new TextValueParser());
            var snip=service.Prepare(state,doc.Id,1,new Doctracker.Core.Geometry.NormalizedRectangle(zone.X,zone.Y,zone.Width,zone.Height),kind,recognized.Text,"Feuil1","A1","test");
            Check(snip.RawText==recognized.Text && snip.ExtractedValue.Contains("FA-001"),"Validation/anomaly discarded selected source text");
        }
        return "PASS: PDF table points add/drag/delete, preview correction, zoom-stable geometry, no accidental snips, validation/anomaly source text";
    }
}
