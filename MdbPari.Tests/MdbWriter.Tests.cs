using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FactoryPari;
using InterfacesPari;

namespace MdbPari.Tests
{
    [TestFixture]
    public class MdbWriterClass
    {
        [Test]
        public void CheckConsistency_No_Exception()
        {
            var factory = new Factory();
            IPariDatabase database = factory.CreatePariDatabase();

            database.CheckConsistency();

        }

        [Test]
        public void AddFieldsTest()
        {
            var factory = new Factory();
            IPariDatabase database = factory.CreatePariDatabase();
            var nrOfFields = database.CheckExistingFields();
        }

    }
}
