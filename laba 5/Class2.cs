using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace laba_5
{
    class Account
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public DateTime OpenDate { get; set; }

        public override string ToString()
        {
            return $"ID: {ID}, Имя: {Name}, Дата открытия: {OpenDate.ToShortDateString()}";
        }
    }
}
