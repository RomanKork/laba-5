using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace laba_5
{
    class Currency
    {
        public string LetterCode { get; set; }
        public int ID { get; set; }
        public string Name { get; set; }

        public override string ToString()
        {
            return $"ID: {ID}, Буквенный код: {LetterCode}, Наименование: {Name}";
        }
    }
}
