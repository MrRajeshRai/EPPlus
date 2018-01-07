using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExamples
{
   
    public static class CustomerEntity
    {
        public static List<Customer> GetCustomers()
        {
            List<Customer> customerList = new List<Customer>
            {
                new Customer {CustomerName = "Customer 1", City = "Mumbai" },
                new Customer {CustomerName = "Customer 2", City = "Redmond"},
                new Customer {CustomerName = "Customer 3", City = "Chennai" },
                new Customer {CustomerName = "Customer 4", City = "Bangalore"},
                new Customer {CustomerName = "Customer 5", City = "Hyderabad"}
            };

            return customerList;
        }
    }

    public class Customer
    {
        [DisplayName("Customer Name")]
        public string CustomerName { get; set; }

        public string City { get; set; }        
    }
}
