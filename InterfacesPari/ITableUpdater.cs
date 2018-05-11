using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InterfacesPari
{
    public interface ITableUpdater
    {
        List<IKategorieRecord> NewKats { get; }
        List<IKategorieRecord> DelKats { get; }
        List<IKategorieRecord> UpdKats { get; }
        List<IRaumRecord> UpdRaume { get; }
        List<IRaumRecord> DelRaume { get; }
        List<IRaumRecord> NewRaume { get; }
        List<IWohnungRecord> WohnungRecords { get; }
    }
}
