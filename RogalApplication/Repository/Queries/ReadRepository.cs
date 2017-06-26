using System.Collections.Generic;
using System.Linq;
using BazaKlientów;
using RogalApplication.Model;
using RogalApplication.Repository.Queries.Interfaces;

namespace RogalApplication.Repository.Queries
{
    public class ReadRepository<T> : IReadRepository<T> where T : Entity
    {
        private readonly CustomerDatabaseContext _context;

        public ReadRepository(CustomerDatabaseContext context)
        {
            _context = context;
        }
        public IList<T> GetAll()
        {
            return _context.Set<T>().ToList();
        }

        public T GetById(int id)
        {
            return _context.Set<T>().FirstOrDefault(x => x.ID == id);
        }
    }
}
