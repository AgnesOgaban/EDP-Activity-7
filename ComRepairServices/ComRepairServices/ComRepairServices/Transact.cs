using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComRepairServices
{
    public class Transact
    {
        public int TransactionId { get; set; }
        public string CustomerId { get; set; }
        public string EmployeeId { get; set; }
        public DateTime TransactionDate { get; set; }
        public decimal TotalAmount { get; set; }
        public string Description { get; set; }
    }
}
