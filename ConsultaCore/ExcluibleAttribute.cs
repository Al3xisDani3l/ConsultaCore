using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsultaCore
{   
    [AttributeUsage(AttributeTargets.Property)]
   
    public class ExcluibleAttribute:Attribute
    {

        private bool _IsExcluible = false;

        public bool IsExcluible { get => _IsExcluible; set => _IsExcluible = value; }
        /// <summary>
        /// Evita que una propiedad sea incluida, en la base de datos;
        /// </summary>
        public ExcluibleAttribute()
        {
            IsExcluible = true;
        }
        
        public ExcluibleAttribute(bool isExcluible)
        {
            IsExcluible = isExcluible;
        }

      
    }
}
