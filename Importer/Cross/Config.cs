﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Importer.Cross
{
    public class Config
    {
        public int SheetNumber { get; set; }
        public int BrandCol { get; set; } = 1;
        public int CodeCol { get; set; } = 2;
        public int DestBrandCol { get; set; } = 3;
        public int DestCodeCol { get; set; } = 4;
    }
}
