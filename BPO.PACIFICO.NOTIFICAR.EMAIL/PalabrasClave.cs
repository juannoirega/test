using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPO.PACIFICO.NOTIFICAR.EMAIL
{
    public class PalabrasClave : IEquatable<PalabrasClave>
    {
        public String clave { get; set; }
        public String palabra { get; set; }

        public bool Equals(PalabrasClave other)
        {
            throw new NotImplementedException();
        }
    }
}
