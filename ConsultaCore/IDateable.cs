using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace ConsultaCore
{
    /// <summary>
    /// Interfaz entre un objeto comun de .Net Framework y Microsoft Access.
    /// Contiene implementaciones que hace posible la interaccion de manera fluida entre un objeto y una Tabla de Access,
    /// Implementando el modelo Propiedad-Campo donde las propiedades representan en una Tabla sus campos.
    /// si se desea que una propiedad no sea incluida en una tabla debe implementar el atributo  <see cref="ExcluibleAttribute"/>
    /// </summary>
   public interface IDateable
    {
        [Excluible]
        int Id { get; set; }
        [Excluible]
        string Tabla { get;}
        [Excluible]
        List<PropertyInfo> Propiedades { get; }
        
    }

    public abstract class IdateableObject : IDateable
    {
        [Excluible]
        public virtual int Id { get; set; }
        [Excluible]
        public virtual string Tabla { get { return GetType().Name;} }
        [Excluible]
        public virtual List<PropertyInfo> Propiedades { get
            {
                return GetType().GetProperties().ToList().Where(o => Attribute.GetCustomAttribute(o, typeof(ExcluibleAttribute)) == null).ToList();
            } }
    }
}
