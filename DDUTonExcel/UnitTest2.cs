using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DDUTonExcel
{
    [TestClass]
    public class UnitTest2
    {
        [TestMethod]
        public void TestMethod1()
        {
        }
    }

    public class BankAccount
    {
        private string custName;
        private double bal;

        public BankAccount(string customerName, double balance)
        {
            custName = customerName;
            bal = balance;
        }

        public double Balance
        { get { return bal; } }

        public void Debit(double amount)
        {
            if (amount < 0)
                throw new ArgumentOutOfRangeException("amount");
            bal -= amount;
        }
    }
}
