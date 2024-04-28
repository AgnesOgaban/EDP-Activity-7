using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComRepairServices
{
    // Class to represent a repair request
    public class RepairRequest
    {
        public int RequestId { get; set; }
        public string CustomerName { get; set; }
        public string DeviceBrand { get; set; }
        public string DeviceModel { get; set; }
        public string IssueDescription { get; set; }
        public DateTime RequestDate { get; set; }
        public string Status { get; set; }
    }


}
