using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPO.PACIFICO
{
    public class Puntuacion : IEquatable<Puntuacion>
    {
        public int contador { get; set; }
        public string palabra { get; set; }


        public bool Equals(Puntuacion other)
        {
            throw new NotImplementedException();
        }
    }
}

