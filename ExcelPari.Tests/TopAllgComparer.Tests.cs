using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelPari;

namespace ExcelPari.Tests
{
    [TestFixture]
    public class TopAllgComparerTests
    {
        [TestCase("ALLG", "ALLG")]
        public void Compare_XSmallerY_ReturnMinus(string x, string y)
        {
            // Arrange
            //var x = "Top A2";
            //var y = "Top A11";
            var lst = new List<string>() { y, x };
            var comparer = new TopAllgComparer();

            // Act 
            lst.Sort(comparer);

            // Assert
            Assert.That(lst.IndexOf(x) == 0);
        }

        [Test]
        public void Compare_List_Ok()
        {
            // Arrange
            //var x = "Top A2";
            //var y = "Top A11";
            var lst = new List<string>() { "ALLG", "TOP AALLG", "TOP AALLG", "TOP AALLG", "TOP AALLG", "TOP AALLG", "ALLG", "TOP AALLG", "TOP AALLG", "ALLG", "TOP AALLG", "TOP AALLG" };
            var comparer = new TopAllgComparer();

            // Act 
            lst.Sort(comparer);

            // Assert
            Assert.That(lst.IndexOf("TOP AALLG") == 0);
        }

    }
}
