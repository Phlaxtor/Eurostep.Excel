namespace Eurostep.Excel.Test;

[TestClass]
public sealed class FunctionTest
{
    [TestMethod]
    public void TestColumnId_Id_to_Name_to_Id()
    {
        for (uint i = 0; i <= 16384; i++)
        {
            ColumnId from = new ColumnId(i);
            ColumnId to = new ColumnId(from.Name);
            Assert.AreEqual(i, to.No);
            Assert.AreEqual(from.ColumnName, to.ColumnName);
        }
    }
}