using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BazaKlientów.Model;

namespace BazaKlientów.Repository.Commands.Interfaces
{
    public interface IWriteRepository<T> where T:Entity
    {
        void Create(T entity);
        void Delete(T entity);
        void Edit(T entity, T updated);
    }
}
