using System;
using System.Collections.Generic;
namespace InterfacesPari
{
    public interface ITableBuilder
    {
        void Build(System.Collections.Generic.List<InterfacesPari.IBlockInfo> blockInfos, List<IWohnungInfo> wohnungInfos);
        System.Collections.Generic.Dictionary<string, InterfacesPari.IKategorieRecord> KatDict { get; set; }
        System.Collections.Generic.List<InterfacesPari.IRaumRecord> RaumTable { get; set; }
        System.Collections.Generic.List<InterfacesPari.IWohnungRecord> WohnungTable { get; set; }
    }
}
