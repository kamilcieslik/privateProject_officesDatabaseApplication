using BazaKlientów;
using RogalApplication.Model;
using RogalApplication.Repository.Commands.Interfaces;

namespace RogalApplication.Repository.Commands
{
    public class WriteRepository<T> : IWriteRepository<T> where T : Entity
    {
        private readonly CustomerDatabaseContext _context;
        public WriteRepository(CustomerDatabaseContext context)
        {
            _context = context;
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
            _context.Entry(entity).CurrentValues.SetValues(updated);
            _context.Entry(entity).Property(o => o.ID).IsModified = false;
            _context.SaveChanges();
        }
    }
}
