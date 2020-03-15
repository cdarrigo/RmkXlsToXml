using System;

namespace RmkXlsToXml
{
    public class RemarketingData
    {
        public string AccountNumber { get; set; }
        public string LoanNumber { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public decimal Balance { get; set; }
        public string Year { get; set; }
        public string Make { get; set; }
        public string Model { get; set; }
        public string Vin { get; set; }
        public int Mileage { get; set; }
        public string RepoAgentName { get; set; }
        public string RepoAgentsLookup { get; set; }
        public string LocationOfUnit { get; set; }
        public DateTime DateOfRepo { get; set; }
        public DateTime DateOfClear { get; set; }

        /// <summary>
        /// FullName is a concatenation of First and Last Name.
        /// </summary>
        public string FullName => $"{FirstName} {LastName}";

    }
}