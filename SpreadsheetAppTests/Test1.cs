using Microsoft.VisualStudio.TestTools.UnitTesting;
using SpreadsheetApp11.Core;
using SpreadsheetApp11;

namespace SpreadsheetApp11.Tests
{
    [TestClass]
    public class SpreadsheetTests
    {
        [TestMethod]
        public void Test_BuiltInFunctions()
        {
            var service = new SpreadsheetService();

            service.SetCell(0, 0, "=min(10,5)");
            service.SetCell(0, 1, "=max(10,5)");
            service.SetCell(0, 2, "=inc(5)");
            service.SetCell(0, 3, "=dec(5)");
            service.SetCell(0, 4, "=min(2+3,4*2)");

            Assert.AreEqual("5", service.GetCellDisplay(0, 0, DisplayMode.Values));
            Assert.AreEqual("10", service.GetCellDisplay(0, 1, DisplayMode.Values));
            Assert.AreEqual("6", service.GetCellDisplay(0, 2, DisplayMode.Values));
            Assert.AreEqual("4", service.GetCellDisplay(0, 3, DisplayMode.Values));
            Assert.AreEqual("5", service.GetCellDisplay(0, 4, DisplayMode.Values));
        }

        [TestMethod]
        public void Test_ComplexCycles()
        {
            var service = new SpreadsheetService();

            // A1 -> B1 -> C1 -> A1 (трикомірковий цикл)
            service.SetCell(0, 0, "=B1");
            service.SetCell(0, 1, "=C1");
            service.SetCell(0, 2, "=A1");

            // Всі мають показувати #ERR
            Assert.AreEqual("#ERR", service.GetCellDisplay(0, 0, DisplayMode.Values));
            Assert.AreEqual("#ERR", service.GetCellDisplay(0, 1, DisplayMode.Values));
            Assert.AreEqual("#ERR", service.GetCellDisplay(0, 2, DisplayMode.Values));
        }

        [TestMethod]
        public void Test_DisplayModes()
        {
            var service = new SpreadsheetService();

            service.SetCell(0, 0, "42");           // Число
            service.SetCell(0, 1, "Hello");        // Текст
            service.SetCell(0, 2, "=10+5");        // Формула

            // Режим значень
            Assert.AreEqual("42", service.GetCellDisplay(0, 0, DisplayMode.Values));
            Assert.AreEqual("Hello", service.GetCellDisplay(0, 1, DisplayMode.Values));
            Assert.AreEqual("15", service.GetCellDisplay(0, 2, DisplayMode.Values));

            // Режим формул
            Assert.AreEqual("42", service.GetCellDisplay(0, 0, DisplayMode.Formulas));
            Assert.AreEqual("Hello", service.GetCellDisplay(0, 1, DisplayMode.Formulas));
            Assert.AreEqual("=10+5", service.GetCellDisplay(0, 2, DisplayMode.Formulas));
        }
    }
}