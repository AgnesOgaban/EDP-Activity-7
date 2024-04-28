using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComRepairServices
{
    // Class to represent a repair request
    public class RepairStatusUpdate
    {
        public int RequestId { get; set; }
        public string NewStatus { get; set; }
        public DateTime UpdateDate { get; set; }
    }
}

