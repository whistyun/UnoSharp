using NUnit.Framework;
using UnoSharp;
using System;

namespace UnoSharp.Tests
{
    public class CellAddressUtilTests
    {
        [Test]
        public void ConvertDateToValueTest()
        {
            var nullDate = new DateTime(1899, 12, 30);

            var inputs = new[] {
                new DateTime(1899,12,28, 17,59,50),
                new DateTime(1899,12,29),
                new DateTime(1900,12,31),
                new DateTime(1904,10,10),
                new DateTime(1990,12, 3),
                new DateTime(1990,12, 3, 17,59,50),
            };

            var expecteds = new[] {
                -1.25011574074074,
                -1,
                366,
                1745,
                33210,
                33210.7498842593
            };


            for (var idx = 0; idx < inputs.Length; ++idx)
            {
                var input = inputs[idx];
                var expected = expecteds[idx];

                Assert.AreEqual(expected.ToString(), Utils.ConvertDateToValue(nullDate, input).ToString());
            }
        }

        [Test]
        public void ConvertValueToDateTest()
        {
            var nullDate = new DateTime(1899, 12, 30);

            var expecteds = new[] {
                new DateTime(1899,12,28, 17,59,50),
                new DateTime(1899,12,29),
                new DateTime(1900,12,31),
                new DateTime(1904,10,10),
                new DateTime(1990,12, 3),
                new DateTime(1990,12, 3, 17,59,50),
            };

            var inputs = new[] {
                -1.25011574074074,
                -1,
                366,
                1745,
                33210,
                33210.7498842593
            };


            for (var idx = 0; idx < inputs.Length; ++idx)
            {
                var input = inputs[idx];
                var expected = expecteds[idx];

                Assert.AreEqual(expected, Utils.ConvertValueToDate(nullDate, input));
            }
        }

        [Test]
        public void CovertIndexToColumnLabelTest()
        {
            int[] inputs = {
                0,
                25,
                26,
                51,
                52,
                701,
                702

            };
            string[] expecteds = {
                "A",
                "Z",
                "AA",
                "AZ",
                "BA",
                "ZZ",
                "AAA"
            };

            for (var idx = 0; idx < inputs.Length; ++idx)
            {
                var input = inputs[idx];
                var expected = expecteds[idx];

                Assert.AreEqual(expected, Utils.CovertIndexToColumnLabel(input));
            }
        }

        [Test]
        public void ConvertColumnLabelToIndexTest()
        {
            int[] expecteds = {
                0,
                25,
                26,
                51,
                52,
                701,
                702

            };
            string[] inputs = {
                "A",
                "Z",
                "AA",
                "AZ",
                "BA",
                "ZZ",
                "AAA"
            };

            for (var idx = 0; idx < inputs.Length; ++idx)
            {
                var input = inputs[idx];
                var expected = expecteds[idx];

                Assert.AreEqual(expected, Utils.ConvertColumnLabelToIndex(input));
            }
        }
    }
}