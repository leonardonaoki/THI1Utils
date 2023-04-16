using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SubstituidorIDDatabasexIdGene
{
    public class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var dicRetorno = ObterDadosExcel(args[0]);
                ProcessarTXT(dicRetorno, args[1], args[2]);
            }
            catch (Exception e)
            {
                Console.WriteLine(@"Por favor preencha corretamente os parâmetros separados por espaços. 
                Lembrando, 1º Parametro = Caminho do Excel em que a coluna L = 'ID Database' e a coluna N = 'ID Gene'
                            2º Parametro = MultiFasta de CDS em formato TXT.
                            3º Parametro = Caminho do Arquivo de saída, exemplo 'C:\Users\Teste\Desktop\result.txt' ");
                Console.WriteLine("Erro Original: " + e);
            }
        }

        private static void ProcessarTXT(Dictionary<string, string> dicRetorno, string multiFastaCDS, string outputFile)
        {
            var linhasExcel = File.ReadAllLines(multiFastaCDS);
            StringBuilder sb = new StringBuilder();
            foreach (var linha in linhasExcel)
            {
                if (linha.StartsWith(">"))
                {
                    int indexBarrinha = linha.IndexOf("|");
                    string descricao = linha.Remove(indexBarrinha).Replace(">", "");

                    string novaLinha;

                    if (dicRetorno.TryGetValue(descricao.ToUpper(), out string valueSaida))
                        novaLinha = ">" + valueSaida;
                    else
                        novaLinha = linha + " - Não foi identificado no excel recebido";

                    sb.AppendLine(novaLinha);
                }
                else
                    sb.AppendLine(linha);
            }

            File.WriteAllText(outputFile, sb.ToString());
        }
        /// <summary>
        /// Método para obter um dicionário de chave-valor entre IdDatabase x IDGene
        /// </summary>
        /// <param name="Excel">Caminho do excel utilizado</param>
        /// <returns></returns>
        private static Dictionary<string, string> ObterDadosExcel(string Excel)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Dictionary<string, string> dicExcel = new Dictionary<string, string>();
            using (var excel = new ExcelPackage())
            {
                using (var stream = File.OpenRead(Excel))
                    excel.Load(stream);

                var pagina1WS = excel.Workbook.Worksheets[0];

                int contador = 2;
                foreach (var cell in pagina1WS.Cells["L2:L301"])
                {
                    string idDatabase = cell.Value?.ToString();
                    string idGene = pagina1WS.Cells[$"N{contador}"].Value?.ToString();
                    contador++;

                    if (!string.IsNullOrWhiteSpace(idDatabase) && !dicExcel.ContainsKey(idDatabase))
                        dicExcel.Add(idDatabase.ToUpper(), idGene);
                }
            }
            return dicExcel;
        }
    }


}

