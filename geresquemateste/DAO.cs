using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using ExcelDataReader;


namespace geresquemateste
{
    public static class DAO
    {
        public static List<CasoDeTeste> ObterDadosDaPlanilha(string caminhoArquivo, bool primeiraLinhaCabecalho = true)
        {
            try
            {
                var casos = new List<CasoDeTeste>();
                using (FileStream stream = File.Open(caminhoArquivo, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                    {
                        string identificador = string.Empty;
                        int linha = 0;
                        while (reader.Read())
                        {
                            linha = linha + 1;

                            if (linha == 1 && primeiraLinhaCabecalho)
                                continue;


                            identificador = reader.GetString((int)CasoDeTeste.ColunaPlanilha.Id);
                            if (String.IsNullOrWhiteSpace(identificador))
                                break;

                            var casoTeste = new CasoDeTeste();
                            casoTeste.Identificador = identificador;
                            casoTeste.Grupo = reader.GetString((int)CasoDeTeste.ColunaPlanilha.Grupo);
                            casoTeste.Nome = reader.GetString((int)CasoDeTeste.ColunaPlanilha.Nome);
                            casos.Add(casoTeste);
                        }
                    }
                }
                return casos;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



    }
}
