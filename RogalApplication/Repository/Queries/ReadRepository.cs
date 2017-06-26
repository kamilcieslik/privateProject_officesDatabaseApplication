using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BazaKlientów.Model;
using BazaKlientów.Repository.Queries.Interfaces;

namespace BazaKlientów.Repository.Queries
{
    public class ReadRepository<T> : IReadRepository<T> where T : Entity
    {
        private readonly CustomerDatabaseContext _context;

        public ReadRepository(CustomerDatabaseContext context)
        {
            this._context = context;
        }
        public IList<T> GetAll()
        {
            return _context.Set<T>().ToList();
        }

        public T GetById(int id)
        {
            return _context.Set<T>().Where(x => x.ID == id).FirstOrDefault();
        }
    }
}
