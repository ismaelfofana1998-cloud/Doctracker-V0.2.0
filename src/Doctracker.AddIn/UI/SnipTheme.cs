using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Doctracker.Core.Models;

namespace Doctracker.AddIn.UI
{
    internal static class SnipTheme
    {
        public static readonly Color Ink = Color.FromArgb(30, 41, 59);
        public static readonly Color Muted = Color.FromArgb(100, 116, 139);
        public static readonly Color Surface = Color.FromArgb(245, 247, 251);
        public static Color ColorFor(SnipType type)
        {
            switch (type)
            {
                case SnipType.Number: return Color.FromArgb(8, 128, 145);
                case SnipType.Date: return Color.FromArgb(79, 70, 190);
                case SnipType.Sum: return Color.FromArgb(173, 128, 0);
                case SnipType.Table: return Color.FromArgb(147, 51, 180);
                case SnipType.Validation: return Color.FromArgb(22, 139, 83);
                case SnipType.Exception: return Color.FromArgb(211, 56, 72);
                default: return Color.FromArgb(44, 110, 219);
            }
        }
        public static string LabelFor(SnipType type)
        {
            switch (type)
            {
                case SnipType.Text: return "Texte";
                case SnipType.Number: return "Nombre";
                case SnipType.Date: return "Date";
                case SnipType.Sum: return "Somme";
                case SnipType.Table: return "Tableau";
                case SnipType.Validation: return "Validation";
                default: return "Anomalie";
            }
        }
        public static Color Tint(SnipType type)
        {
            var c = ColorFor(type);
            return Color.FromArgb(225 + c.R * 30 / 255, 225 + c.G * 30 / 255, 225 + c.B * 30 / 255);
        }
        // Original vector glyphs; shared between Office, the pane and document tools.
        public static Bitmap Icon(string key, int size = 24, Color? color = null)
        {
            var image = new Bitmap(size, size);
            using (var g = Graphics.FromImage(image))
            using (var pen = new Pen(color ?? Ink, 1.7f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.ScaleTransform(size / 24f, size / 24f);
                switch (key)
                {
                    case "CrossReference":
                        g.DrawEllipse(pen,3,5,11,7);g.DrawEllipse(pen,10,12,11,7);g.DrawLine(pen,9,9,15,15);break;
                    case "ExportPdf":
                        g.DrawRectangle(pen,4,3,11,18);g.DrawLine(pen,10,12,22,12);g.DrawLines(pen,new[]{new PointF(18,8),new PointF(22,12),new PointF(18,16)});break;
                    case "Comment":
                        g.DrawRectangle(pen,3,4,18,13);g.DrawLines(pen,new[]{new PointF(7,17),new PointF(7,21),new PointF(12,17)});g.DrawLine(pen,7,8,17,8);g.DrawLine(pen,7,12,14,12);break;
                    case "Delete":
                        g.DrawLine(pen,4,6,20,6);g.DrawRectangle(pen,6,6,12,15);g.DrawLine(pen,9,3,15,3);g.DrawLine(pen,10,10,10,17);g.DrawLine(pen,14,10,14,17);break;
                    case "Validation": case "ReviewProof":
                        g.DrawLines(pen, new[] {new PointF(4,12),new PointF(9,17),new PointF(20,6)}); break;
                    case "Exception":
                        g.DrawLine(pen,6,6,18,18); g.DrawLine(pen,18,6,6,18); break;
                    case "Text":
                        g.DrawLine(pen,5,5,19,5); g.DrawLine(pen,12,5,12,19); g.DrawLine(pen,8,19,16,19); break;
                    case "Number":
                        g.DrawLine(pen,10,4,7,20); g.DrawLine(pen,17,4,14,20); g.DrawLine(pen,5,9,20,9); g.DrawLine(pen,4,15,19,15); break;
                    case "Sum":
                        g.DrawLines(pen,new[]{new PointF(19,5),new PointF(6,5),new PointF(13,12),new PointF(6,19),new PointF(19,19)}); break;
                    case "Date": case "Table":
                        g.DrawRectangle(pen,4,5,16,15); g.DrawLine(pen,4,10,20,10);
                        if(key=="Date") { g.DrawLine(pen,8,3,8,7); g.DrawLine(pen,16,3,16,7); g.DrawRectangle(pen,8,13,3,3); }
                        else {g.DrawLine(pen,4,15,20,15);g.DrawLine(pen,10,5,10,20);g.DrawLine(pen,15,5,15,20);} break;
                    case "OcrDocuments":case "SearchDocuments":
                        g.DrawEllipse(pen,4,3,12,12);g.DrawLine(pen,15,15,21,21);break;
                    case "ImportDocuments": case "Plus":
                        g.DrawLine(pen,12,5,12,19);g.DrawLine(pen,5,12,19,12);break;
                    case "Minus":g.DrawLine(pen,5,12,19,12);break;
                    case "Previous":g.DrawLines(pen,new[]{new PointF(15,5),new PointF(8,12),new PointF(15,19)});break;
                    case "Next":g.DrawLines(pen,new[]{new PointF(9,5),new PointF(16,12),new PointF(9,19)});break;
                    case "More":
                        using(var b=new SolidBrush(pen.Color)) for(var x=5;x<=19;x+=7)g.FillEllipse(b,x-1,11,2,2);break;
                    case "Fit":
                        g.DrawRectangle(pen,6,3,12,18);g.DrawLine(pen,2,9,2,15);g.DrawLine(pen,22,9,22,15);break;
                    case "Match":case "SetMatchInput":case "SetMatchOutput":
                        g.DrawRectangle(pen,3,4,6,6);g.DrawRectangle(pen,15,14,6,6);g.DrawLines(pen,new[]{new PointF(12,7),new PointF(18,7),new PointF(18,11)});g.DrawLines(pen,new[]{new PointF(12,17),new PointF(6,17),new PointF(6,13)});break;
                    default:
                        g.DrawLines(pen,new[]{new PointF(14,3),new PointF(5,3),new PointF(5,21),new PointF(19,21),new PointF(19,8),new PointF(14,3),new PointF(14,8),new PointF(19,8)});
                        g.DrawLine(pen,9,12,15,12);g.DrawLine(pen,9,16,15,16);break;
                }
            }
            return image;
        }
        public static Button Button(string text, string icon = null, Color? color = null)
        {
            var button = new Button { Text = text, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlatStyle = FlatStyle.Flat, BackColor = Color.White, ForeColor = color ?? Ink,
                Padding = new Padding(6, 4, 6, 4), Margin = new Padding(2), Cursor = Cursors.Hand,
                TextImageRelation = TextImageRelation.ImageBeforeText, AccessibleName = text, UseVisualStyleBackColor = false };
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = Surface;
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(225,232,244);
            if(icon != null)
            {
                Action refresh = () => { var old=button.Image; button.Image=Icon(icon,Math.Max(20,(int)Math.Ceiling(button.Font.Height*1.25)),color); old?.Dispose(); };
                refresh(); button.FontChanged += (s,e)=>refresh();
                button.Disposed += (s,e)=>button.Image.Dispose();
            }
            return button;
        }
    }
}
