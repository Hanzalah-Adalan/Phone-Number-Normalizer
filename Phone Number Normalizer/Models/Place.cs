using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Phone_Number_Normalizer.Models
{
    public class Place
    {
        public string State { get; set; }
        public string District { get; set; }
        public string GroupKey { get; set; }
        public string Address { get; set; }

        public List<Place> Children { get; set; } = new List<Place>();
        public int DuplicateCount { get; set; }
    }
}
