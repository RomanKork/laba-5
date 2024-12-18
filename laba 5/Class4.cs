using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace laba_5
{
    class CurrencyRate
    {
        public int ID { get; set; }
        public int CurrencyID { get; set; }
        public DateTime Date { get; set; }
        public decimal Rate { get; set; }

        public override string ToString()
        {
            return $"ID: {ID}, Валюта ID: {CurrencyID}, Дата: {Date.ToShortDateString()}, Курс: {Rate}";
        }
    }
}
