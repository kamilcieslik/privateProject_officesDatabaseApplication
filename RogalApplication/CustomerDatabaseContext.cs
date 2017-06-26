using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RogalApplication.Model;

namespace BazaKlientów
{
    public class CustomerDatabaseContext : DbContext
    {

       
        public virtual DbSet<Customer> Customers { get; set; }

        public CustomerDatabaseContext(): base("name=CustomerDatabaseContext")
        {
            
        }
    }
}
