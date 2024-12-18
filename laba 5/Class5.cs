using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace laba_5
{
    class Transaction
    {
        public int ID { get; set; }
        public int AccountID { get; set; }
        public int CurrencyID { get; set; }
        public DateTime Date { get; set; }
        public decimal Amount { get; set; }

        public override string ToString()
        {
            return $"ID: {ID}, Счет ID: {AccountID}, Валюта ID: {CurrencyID}, Дата: {Date.ToShortDateString()}, Сумма: {Amount}";
        }
    }

}
