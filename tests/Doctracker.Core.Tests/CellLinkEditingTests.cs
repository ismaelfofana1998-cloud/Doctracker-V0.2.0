using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Doctracker.Core.Models;
using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class CellLinkEditingTests : IDisposable
    {
        private readonly string root=Path.Combine(Path.GetTempPath(),"doctracker-links-"+Guid.NewGuid().ToString("N"));
        private static CellLinkRecord Link(string address,params string[] ids) => new CellLinkRecord {
            WorksheetName="Factures",WorksheetCodeName="Sheet1",CellAddress=address,SnipIds=ids.ToList()};
        [Fact] public void Range_deletion_preserves_shared_proofs_and_unrelated_cells()
        {
            var store=new ProjectStore(root);var state=new ProjectState();
            state.Snips.AddRange(new[]{new SnipRecord {Id="one"},new SnipRecord {Id="shared"},new SnipRecord {Id="outside"}});
            var selected=Link("A1","one","shared");var other=Link("B2","shared","outside");
            state.CellLinks.AddRange(new[]{selected,other});
            Assert.Equal(1,new SnipService(store,new TextValueParser()).DeleteCellLinks(state,new[]{selected},"test"));
            Assert.Equal(new[]{"shared","outside"},state.Snips.Select(s=>s.Id));
            Assert.Same(other,Assert.Single(state.CellLinks));
            var loaded=new ProjectStore(root).LoadOrCreate("");
            Assert.Equal("Sheet1",Assert.Single(loaded.CellLinks).WorksheetCodeName);
            Assert.Equal(2,loaded.Snips.Count);
        }
        [Fact] public void Failed_range_deletion_restores_all_metadata_and_audit()
        {
            var store=new ProjectStore(root);var state=new ProjectState();var snip=new SnipRecord {Id="one"};
            state.Snips.Add(snip);var link=Link("A1",snip.Id);state.CellLinks.Add(link);
            Directory.CreateDirectory(store.MetadataPath);
            Assert.ThrowsAny<Exception>(()=>new SnipService(store,new TextValueParser()).DeleteCellLinks(state,new[]{link},"test"));
            Assert.Same(snip,Assert.Single(state.Snips));Assert.Same(link,Assert.Single(state.CellLinks));Assert.Empty(state.AuditTrail);
        }
        [Fact] public void Selection_deletion_uses_stable_sheet_identity_not_display_name()
        {
            var state=new ProjectState();state.Snips.Add(new SnipRecord {Id="one"});state.CellLinks.Add(Link("A1","one"));
            var selected=Link("A1","one");selected.WorksheetName="Nouveau nom";
            new SnipService(new ProjectStore(root),new TextValueParser()).DeleteCellLinks(state,new[]{selected},"test");
            Assert.Empty(state.CellLinks);Assert.Empty(state.Snips);
        }
        [Fact] public void Single_snip_deletion_preserves_other_links_sheet_identity()
        {
            var state=new ProjectState();state.Snips.AddRange(new[]{new SnipRecord {Id="one"},new SnipRecord {Id="two"}});
            state.CellLinks.Add(Link("A1","one","two"));
            new SnipService(new ProjectStore(root),new TextValueParser()).Delete(state,"one","test");
            Assert.Equal("Sheet1",Assert.Single(state.CellLinks).WorksheetCodeName);
            Assert.Equal("two",Assert.Single(state.CellLinks[0].SnipIds));
        }
        public void Dispose(){if(Directory.Exists(root))Directory.Delete(root,true);}
    }
}
