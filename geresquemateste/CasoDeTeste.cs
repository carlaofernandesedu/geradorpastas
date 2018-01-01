using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace geresquemateste
{
    public class CasoDeTeste
    {
        public string Identificador { get; set; }
        public string Grupo { get; set; }
        public string Nome { get; set; }

        public enum ColunaPlanilha
        {
            Id = 0,
            Grupo =4,
            Nome = 5
        }

        public string NomeFormatado => String.Format("[{0} - {1}]", Identificador, Nome);
    }
}
