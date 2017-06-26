using System.Collections.Generic;
using RogalApplication.Model;

namespace RogalApplication.Repository.Queries.Interfaces
{
    public interface IReadRepository<T> where T : Entity
    {
        IList<T> GetAll();
        T GetById(int id);
    }
}
