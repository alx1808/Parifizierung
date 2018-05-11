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
    public class TextNumSortComparerTests
    {
        [TestCase("a", "b")]
        [TestCase("1", "2")]
        [TestCase("2", "12")]
        [TestCase("Top 2", "Top 12")]
        [TestCase("Top 2 a", "Top 12 a")]
        [TestCase("Top 2 a", "Top 2 b")]
        [TestCase("2 a", "12 a")]
        [TestCase("67", "Top 1")]
        [TestCase("222222222222222222222222", "1111111111111111111111111")]
        [TestCase("1111111111111111111111111", "2222222222222222222222222")]
        public void Compare_XSmallerY_ReturnMinus(string x, string y)
        {
            // Arrange
            //var x = "Top A2";
            //var y = "Top A11";
            var lst = new List<string>() { y, x };
            var comparer = new TextNumSortComparer();

            // Act 
            lst.Sort(comparer);

            // Assert
            Assert.That(lst.IndexOf(x) == 0);
        }
    }
}
