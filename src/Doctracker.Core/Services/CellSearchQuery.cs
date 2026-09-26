namespace Doctracker.Core.Services
{
    // A manual query belongs to the current cell, never to the next selected cell.
    public sealed class CellSearchQuery
    {
        private string cellKey,cellText;
        public string Text {get;private set;}="";
        public long Revision {get;private set;}
        public void FollowCell(string key,string value)
        {
            value=value??"";
            if(key==cellKey && value==cellText)return;
            cellKey=key;cellText=value;Text=value;Revision++;
        }
        public void Edit(string value)
        {value=value??"";if(Text==value)return;Text=value;Revision++;}
    }
}
