using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BazaKlientów.Model;

namespace BazaKlientów.Repository.Queries.Interfaces
{
    public interface IReadRepository<T> where T : Entity
    {
        IList<T> GetAll();
        T GetById(int id);
    }
}
