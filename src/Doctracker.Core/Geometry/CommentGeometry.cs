using System;

namespace Doctracker.Core.Geometry
{
    public static class CommentGeometry
    {
        // 0 move, 1 NW, 2 N, 3 NE, 4 E, 5 SE, 6 S, 7 SW, 8 W.
        public static NormalizedRectangle Transform(NormalizedRectangle original,double dx,double dy,int handle)
        {
            if(double.IsNaN(dx)||double.IsInfinity(dx)||double.IsNaN(dy)||double.IsInfinity(dy))throw new ArgumentOutOfRangeException();
            if(handle<0||handle>8)throw new ArgumentOutOfRangeException(nameof(handle));
            if(handle==0)return new NormalizedRectangle(Clamp(original.X+dx,0,1-original.Width),Clamp(original.Y+dy,0,1-original.Height),original.Width,original.Height);
            var left=original.X;var top=original.Y;var right=left+original.Width;var bottom=top+original.Height;
            var minWidth=Math.Min(.03,original.Width);var minHeight=Math.Min(.015,original.Height);
            if(handle==1||handle==7||handle==8)left=Clamp(left+dx,0,right-minWidth);
            if(handle==3||handle==4||handle==5)right=Clamp(right+dx,left+minWidth,1);
            if(handle==1||handle==2||handle==3)top=Clamp(top+dy,0,bottom-minHeight);
            if(handle==5||handle==6||handle==7)bottom=Clamp(bottom+dy,top+minHeight,1);
            return new NormalizedRectangle(left,top,right-left,bottom-top);
        }
        private static double Clamp(double value,double min,double max)=>Math.Max(min,Math.Min(max,value));
    }
}
