﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Importer
{
    public enum PriceFormat { Excel, Csv };
    public class Price
    {
        protected string[] _prefixes = null;
        public string Title { get; set; }
        public PriceFormat Format { get; set; }
        public BindingList<Extra> Extras { get; protected set; }
        public string NameCol { get; set; }
        public string PriceCol { get; set; }
        public string VendorCol { get; set; }
        public string SheetNumber { get; set; }
        public string BeginWith { get; set; }
        public string AmountCol { get; set; }
        public string AvailabilityCol { get; set; }
        public string CodeCol { get; set; }
        public string Coeficients { get; set; }
        public string Prefixes { get; set; }
        public bool RemoveAllPrefixes { get; set; }
        public string Rate { get; set; }
        public string FileName { get; set; }
        public string DefaultAmount { get; set; }

        public Price()
        {
            Format = PriceFormat.Excel;
            Extras = new BindingList<Extra>();
        }
        public override string ToString()
        {
            return Title;
        }

        public string RemovePrefix(string p)
        {
            var newVal = p.Trim();
            if (this.RemoveAllPrefixes)
            {
                var words = newVal.Split(' ').ToList();
                if (words.Count > 1)
                {
                    words.RemoveAt(0);
                    return String.Join(" ", words);
                }
            }
            else
            {
                foreach (var item in _prefixes)
                {
                    if (_prefixes.Length > 0)
                    {
                        if (newVal.StartsWith(item, StringComparison.OrdinalIgnoreCase))
                        {
                            return newVal.Replace(item, " ").Trim();
                        }
                    }
                }
            }
            return newVal;
        }

        public void Init()
        {
            if (Prefixes == null)
            {
                _prefixes = new string[0];
            }
            else
            {
                _prefixes = Prefixes.Replace("\n", "").Split('\r');
            }
        }

    }
}
