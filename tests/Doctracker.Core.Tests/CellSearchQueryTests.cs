using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class CellSearchQueryTests
    {
        [Fact] public void New_cell_replaces_previous_query_including_manual_searches()
        {
            var query=new CellSearchQuery();query.FollowCell("Sheet1!A1","BL001");query.Edit("manual");
            query.FollowCell("Sheet1!A2","BL002");Assert.Equal("BL002",query.Text);
        }
        [Fact] public void Deferred_selection_callback_keeps_manual_query_on_the_same_cell()
        {
            var query=new CellSearchQuery();query.FollowCell("Sheet1!A1","BL001");query.Edit("001");
            var revision=query.Revision;query.FollowCell("Sheet1!A1","BL001");
            Assert.Equal("001",query.Text);Assert.Equal(revision,query.Revision);
        }
        [Fact] public void Changed_value_or_sheet_refreshes_query_and_invalidates_old_results()
        {
            var query=new CellSearchQuery();query.FollowCell("Sheet1!A1","BL001");var revision=query.Revision;
            query.FollowCell("Sheet1!A1","BL009");Assert.Equal("BL009",query.Text);Assert.True(query.Revision>revision);
            query.Edit("custom");query.FollowCell("Sheet2!A1","BL009");Assert.Equal("BL009",query.Text);
        }
        [Fact] public void Blank_cell_clears_previous_query()
        {
            var query=new CellSearchQuery();query.FollowCell("A1","old");query.FollowCell("A2",null);Assert.Equal("",query.Text);
        }
    }
}
