using CsvHelper.Configuration.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project.Models
{
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string Department { get; set; }

        public double TheoryScore { get; set; }
        public double PracticalScore { get; set; }

        // שדה לחישוב
        [Ignore]
        public double FinalScore { get; set; }

        [Ignore]
        public string Message { get; set; }
    }
}
