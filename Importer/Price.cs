using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Importer
{
    public enum PriceFormat { Excel, Csv, DBF };
    public class Price
    {
        protected string[] _prefixes = null;
        protected string[] _suffixes = null;
        protected Dictionary<String, Double> _brandExtra = new Dictionary<string, double>();

        public string Title { get; set; }
        public PriceFormat Format { get; set; }
        public BindingList<Extra> Extras { get; protected set; }
        public string NameCol { get; set; }
        public string PriceCol { get; set; }
        public string VendorCol { get; set; }
        public string SubPercent { get; set; }
        public string SheetNumber { get; set; }
        public string ProviderCode { get; set; }
        public string BeginWith { get; set; }
        public string AmountCol { get; set; }
        public string AvailabilityCol { get; set; }
        public string CodeCol { get; set; }
        public string Coeficients { get; set; }
        public string Prefixes { get; set; }
        public string Suffixes { get; set; }
        public bool RemoveAllPrefixes { get; set; }
        public bool RoundPrice { get; set; }
        public bool ReplaceBrand { get; set; }
        public string NewBrand { get; set; }
        public string Rate { get; set; }
        public string FileName { get; set; }
        public string DefaultAmount { get; set; }
        public string BrandExtra { get; set; }

        public Price()
        {
            Format = PriceFormat.Excel;
            Extras = new BindingList<Extra>();
        }
        public override string ToString()
        {
            return Title;
        }

        public int getSubPercent()
        {
            if (SubPercent == "")
            {
                return 0;
            }
            int val;
            Int32.TryParse(SubPercent, out val);
            return val;
        }

        public string RemovePrefix(string p)
        {
            var newVal = p.Trim();
            if (this.RemoveAllPrefixes)
            {
                var words = newVal.Split(' ').ToList();
                if (words.Count > 1 && words[0].Length <= 3 && !words[0].Any(Char.IsDigit))
                {
                    words.RemoveAt(0);
                    newVal = String.Join(" ", words);
                }
            }
            if (_prefixes.Length > 0)
            {
                foreach (var item in _prefixes)
                {
                    if ((item.Length > 0) && (newVal.StartsWith(item, StringComparison.OrdinalIgnoreCase)))
                    {
                        newVal = newVal.Replace(item, "").Trim();
                        break;
                    }
                }
            }
            foreach (var item in _suffixes)
            {
                if ((item.Length > 0) && newVal.EndsWith(item, StringComparison.OrdinalIgnoreCase))
                {
                    newVal = ReplaceLastOccurrence(newVal, item, "").Trim();
                    break;
                }
            }
            return newVal;
        }

        public static string ReplaceLastOccurrence(string Source, string Find, string Replace)
        {
            int place = Source.LastIndexOf(Find);

            if (place == -1)
                return Source;

            string result = Source.Remove(place, Find.Length).Insert(place, Replace);
            return result;
        }

        public double GetExtraByBrand(String brand)
        {
            double result = 0;
            var upperBrand = brand.ToUpper();

            if (_brandExtra.ContainsKey(upperBrand))
            {
                result = _brandExtra[upperBrand];
            }

            return result;
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
            if (Suffixes == null)
            {
                _suffixes = new string[0];
            }
            else
            {
                _suffixes = Suffixes.Replace("\n", "").Split('\r');
            }
            if (BrandExtra != null)
            {
                string[] lines;

                if (BrandExtra.IndexOf('\r') > -1)
                {
                    lines = BrandExtra.Replace("\n", "").Split('\r');
                }
                else
                {
                    lines = BrandExtra.Split('\n');
                }

                foreach (var line in lines)
                {
                    var items = line.Split(' ');
                    Double extra = 0;
                    Double.TryParse(items.Last(), out extra);
                    var brand = String.Join(" ", items.Reverse().Skip(1).Reverse());
                    _brandExtra.Add(brand.ToUpper(), extra);
                }
            }
        }

    }
}
