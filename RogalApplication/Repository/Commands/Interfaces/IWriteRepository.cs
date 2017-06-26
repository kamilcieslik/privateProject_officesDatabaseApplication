using RogalApplication.Model;

namespace RogalApplication.Repository.Commands.Interfaces
{
    public interface IWriteRepository<in T> where T:Entity
    {
        void Create(T entity);
        void Delete(T entity);
        void Edit(T entity, T updated);
    }
}
