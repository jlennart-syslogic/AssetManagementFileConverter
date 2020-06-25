using System;
using System.Collections.Generic;
using System.Text;

namespace MPSAssestManagementFileConverter.BusinessLogic
{
    public struct OutputFileRecord
    {
        //OUTPUT CSV
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public string Username { get; set; }
        public string Location { get; set; }
        public string PhoneNumber { get; set; }
        public string JobTitle { get; set; }
        public string EmployeeNumber { get; set; }
        public string Company { get; set; }
    }
}
