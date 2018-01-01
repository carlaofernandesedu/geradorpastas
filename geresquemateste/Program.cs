using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace geresquemateste
{
    class Program
    {
        const string msgParametrosNaoInformados = "Favor informar o caminho completo de localizacao da planilha";
        const string msgParametrosIncorretos = "Parametro(s) Incorreto(s)";
        const string msgPastaNaoEncontrada = "Caminho pasta {0} nao encontrado";
        const string msgCriarEstrutura = "Criando pastas para os casos de teste a partir da pasta {0}";
        const string msgFinalizarEstrutura = "Termino da Execucao";
        const string msgAtualizacaoEstruturaCasoTeste = "Atualizando Grupo:{0} \\ Nome:{1}";
        private static string pastaSistema = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);


        static void Main(string[] args)
        {
            //Caso informando o xlsx no proprio local
            //args = new string[] { "roteirotestes.xlsx" };
            //Caso informado o xlsx em outro local 
            //args = new string[] { "f:\\temp\\roteirotestes.xlsx" };
            //Caso informado o xlsx e pasta base em outro local 
            args = new string[] { "f:\\temp\\roteirotestes.xlsx","f:\\temp\\roteiro" };
            //Caso nao informando parametros 
            //Caso informado parametros incorretos 
            //args = new string[] { "f:\\temp\\xoteirotestes.xlsx" };
            //Caso informado parametros incorretos 
            //args = new string[] { "f:\\temp\\roteirotestes.xlsx","f:\\temp\\xt" };

            try
            {
                Executar(args);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
            Console.ResetColor();
            Console.ReadLine();
        }

        static void Executar(string[] args)
        {
            string planilha = string.Empty;
            string pastaBase = string.Empty;
            string retorno;
            if (args.Any())
                planilha = args[0];

            retorno = ValidarArgumentos(args);
            if (String.IsNullOrEmpty(retorno))
            {
                planilha = File.Exists(planilha) ? planilha : Path.Combine(pastaSistema, planilha);
                pastaBase = args.Count() > 1 ? args[1] : pastaSistema;
                var casos = DAO.ObterDadosDaPlanilha(planilha);
                GerarEstruturaPasta(casos, pastaBase);
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine(retorno);
            }
        }

        static string ValidarArgumentos(string[] args)
        {
            var retorno = string.Empty;

            if (!args.Any())
            {
                retorno = msgParametrosNaoInformados;
            }
            else
            {
                var planilha = args[0];
                var pastaBase = args.Count() > 1 ? args[1] : "";

                if (!(File.Exists(planilha) || File.Exists(Path.Combine(pastaSistema, planilha))))
                    retorno = String.Format(msgPastaNaoEncontrada, planilha);
                else if (!(String.IsNullOrWhiteSpace(pastaBase) || Directory.Exists(pastaBase)))
                    retorno = String.Format(msgPastaNaoEncontrada, pastaBase);
            }
            return retorno;
        }

        static void GerarEstruturaPasta(List<CasoDeTeste> casos, string pastaBase)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(String.Format(msgCriarEstrutura,pastaBase));
            casos.ForEach(item =>
            {
                var subPasta = Path.Combine(pastaBase, item.Grupo);
                var subPastaCaso = Path.Combine(pastaBase, item.Grupo, item.NomeFormatado);
                if (!Directory.Exists(subPasta))
                    Directory.CreateDirectory(subPasta);

                if (!Directory.Exists(subPastaCaso))
                    Directory.CreateDirectory(subPastaCaso);
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine(String.Format(msgAtualizacaoEstruturaCasoTeste, item.Grupo, item.NomeFormatado));
            });
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(msgFinalizarEstrutura);
        }
    }
}
