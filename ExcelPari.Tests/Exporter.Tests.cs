using FactoryPari;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPari.Tests
{
    [TestFixture]
    public class ExporterTests
    {
        [Test]
        public void Export_DatabaseNull_ArgumentNullException()
        {
            // Arrange
            var exporter = new Exporter();
            ArgumentNullException exp = null;

            // Act 
            try
            {
                exporter.ExportNW(null, "", 0);
            }
            catch (Exception ex)
            {
                exp = ex as ArgumentNullException;
            }

            // Assert
            Assert.That(exp != null);
            Assert.That(exp.ParamName == "database");
        }

        //[Test]
        //public void Export_LocationHintHull_ArgumentNullException()
        //{
        //    // Arrange
        //    var exporter = new Exporter();
        //    ArgumentNullException exp = null;
        //    var fact = new Factory();

        //    // Act 
        //    try
        //    {
        //        exporter.Export(fact.CreatePariDatabase(), null, 0);
        //    }
        //    catch (Exception ex)
        //    {
        //        exp = ex as ArgumentNullException;
        //    }

        //    // Assert
        //    Assert.That(exp != null);
        //    Assert.That(exp.ParamName == "locationHint");
        //}

        [Test]
        public void Export_TargetFileExist_ArgumentException()
        {
            // Arrange
            var exporter = new Exporter();
            ArgumentException exp = null;
            var fact = new Factory();
            var targetFile = System.Reflection.Assembly.GetExecutingAssembly().Location;

            // Act 
            try
            {
                exporter.ExportNW(fact.CreatePariDatabase(), targetFile, 0);
            }
            catch (Exception ex)
            {
                exp = ex as ArgumentException;
            }

            // Assert
            Assert.That(exp != null);
            Assert.That(exp.ParamName == "locationHint");
        }

        [Test]
        public void Export_WrongProjektId_InvalidOperationExc()
        {
            // Arrange
            var exporter = new Exporter();
            InvalidOperationException exp = null;
            var fact = new Factory();
            var tmpFile = System.IO.Path.GetTempFileName();
            var targetFile = Path.Combine(Path.GetDirectoryName(tmpFile), Path.GetFileNameWithoutExtension(tmpFile) + ".xlsx");

            int projektId = -1;

            // Act 
            try
            {
                exporter.ExportNW(fact.CreatePariDatabase(), targetFile, projektId);
            }
            catch (Exception ex)
            {
                exp = ex as InvalidOperationException;
            }

            // Assert
            Assert.That(exp != null);
            Assert.That(exp.Message.StartsWith("ProjektInfo with id"));
        }

        [Test]
        public void Export_ValidData_FileExists()
        {

            // Arrange
            var exporter = new Exporter();
            var fact = new Factory();
            var tmpFile = System.IO.Path.GetTempFileName();
            var targetFile = Path.Combine(Path.GetDirectoryName(tmpFile), Path.GetFileNameWithoutExtension(tmpFile) + ".xlsx");
            var database = fact.CreatePariDatabase();
            var projektId = database.ListProjInfos()[0].ProjektId;

            // Act 
            try
            {
                exporter.ExportNW(database, targetFile, projektId);

                // Assert
                Assert.That(File.Exists(targetFile));
            }
            catch (Exception ex)
            {
                var x = ex;
            }
            finally
            {
                if (File.Exists(targetFile)) File.Delete(targetFile);
            }
        }
    }
}
