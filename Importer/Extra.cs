using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Importer
{
    public class Extra
    {
        public Double From { get; set; }
        public Double To { get; set; }
        public Double Percent { get; set; }

        public Boolean isAcceptable(Double price)
        {
            return To > price && price >= From;
        }

        public Double UpdatePrice(Double price)
        {
            return price * Percent / 100;
        }
    }
}
