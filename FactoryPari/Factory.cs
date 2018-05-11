using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FactoryPari
{
    public class Factory : IFactory
    {
        #region log4net Initialization
        private static log4net.ILog log = log4net.LogManager.GetLogger(typeof(Factory));
        static Factory()
        {
            if (log4net.LogManager.GetRepository(System.Reflection.Assembly.GetExecutingAssembly()).Configured == false)
            {
                log4net.Config.XmlConfigurator.ConfigureAndWatch(
                    new System.IO.FileInfo(
                        System.IO.Path.Combine(
                            new System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).DirectoryName,
                            "_log4net.config"
                        )
                    )
                );
            }
        }
        #endregion

        public IProjektInfo CreateProjectInfo()
        {
            return new ProjektInfo();
        }
        public ISubInfo CreateSubInfo()
        {
            return new FactoryPari.ProjektInfo.SubInfo();
        }
        public IPariDatabase CreatePariDatabase()
        {
            log.Debug("CreatePariDatabase");
            return new MdbPari.MdbWriter(this);
        }
        public IKategorieRecord CreateKategorie()
        {
            return new KategorieRecord();
        }
        public IKategorieZaRecord CreateZaKategorie()
        {
            return new KategorieZaRecord();
        }
        public IKategorieRecord CreateKategorie(IRaumRecord raumRecord)
        {
            return new KategorieRecord(raumRecord);
        }

        public IVisualOutput CreateVisualOutputHandler()
        {
            log.Debug("CreateVisualOutputHandler");
            return new ExcelPari.Exporter();
        }

        public IRaumRecord CreateRaumRecord()
        {
            return new RaumRecord();
        }

        public IRaumRecord CreateRaumRecord(IBlockInfo acadBlockInfo)
        {
            var raumRecord = new RaumRecord();
            raumRecord.UpdateValuesFrom(acadBlockInfo);
            return raumRecord;
        }
    }
}
