using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BazaKlientów.Model;
using BazaKlientów.Repository.Commands.Interfaces;

namespace BazaKlientów.Repository.Commands
{
    public class WriteRepository<T> : IWriteRepository<T> where T : Entity
    {
        private readonly CustomerDatabaseContext _context;
        public WriteRepository(CustomerDatabaseContext context)
        {
            this._context = context;
        }

        public void Create(T entity)
        {
            _context.Set<T>().Add(entity);
            _context.SaveChanges();
        }

        public void Delete(T entity)
        {
            _context.Set<T>().Remove(entity);
            _context.SaveChanges();
        }

        public void Edit(T entity, T updated)
        {
            _context.Entry<T>(entity).CurrentValues.SetValues(updated);
            _context.Entry<T>(entity).Property(o => o.ID).IsModified = false;
            _context.SaveChanges();
        }
    }
}
