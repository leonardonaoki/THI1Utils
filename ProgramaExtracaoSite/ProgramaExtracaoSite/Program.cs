using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.RegularExpressions;

namespace ExtracaoMemeParaXML
{
    public class Program
    {
        private static readonly HttpClient httpClient = new HttpClient();

        public static Regex regexUrls = new Regex(@"""url"".+""(.*)""");
        public static Regex matrixDetails = new Regex(@"matrix-detail"">.*</table>\n            <hr>", RegexOptions.Singleline);
        public static Regex Name = new Regex(@"Name:.*\n              <td>(.*)<");
        public static Regex MatrixID = new Regex(@"Matrix ID:.*\n              <td>(.*)<");
        public static Regex Class = new Regex(@"Class:.*\n              <td>(.*)<");
        public static Regex Family = new Regex(@"Family:.*\n              <td>(.*)<");
        public static Regex Collection = new Regex(@"Collection:.*\n              <td>(.*)\n              (.*)");
        public static Regex Taxon = new Regex(@"Taxon:.*\n              <td>(.*)<");
        public static Regex Species = new Regex(@"Species:.*\n              <td>(.*)\n              \n              (.*)\n              (.*)");
        public static Regex DataType = new Regex(@"Data Type:.*\n              <td>(.*)<");
        public static Regex Validation = new Regex(@"Validation:.*\n              .*\n                \n                \n                .*\n\n                (.*)");
        public static Regex UniprotID = new Regex(@"Uniprot ID:.*\n              .*\n              \n              .*""> (.*) <");

        public static void Main(string[] args)
        {
            try
            {
                List<ObjetoMeme> ListaFinal = BaterSite(args[0]);
                GravarDadosExcel(ListaFinal, args[1]);
            }
            catch (Exception e)
            {
                Console.WriteLine(@"Por favor preencha corretamente os parâmetros separados por espaços. Lembrando, 1º Parametro = Link Site,
                2º Parametro = Caminho do Arquivo de saída, exemplo: 'C:\Users\Teste\MEME_INFO.xlsx'");
                Console.WriteLine("Erro Original: " + e);
            }
        }

        private static List<ObjetoMeme> BaterSite(string link)
        {
            var responseSite = httpClient.GetStringAsync(link).Result;

            List<string> Urls = new List<string>();

            var matchesUrls = regexUrls.Matches(responseSite);
            foreach (Match match in matchesUrls)
                Urls.Add(match.Groups[1].Value);

            List<ObjetoMeme> listaRetorno = new List<ObjetoMeme>();
            for (int i = 0; i < Urls.Count; i++)
            {
                ObjetoMeme objetoRetorno = new ObjetoMeme();

                var responseSite2 = httpClient.GetStringAsync(Urls[i]).Result;
                Match matchesMatrixDetails = matrixDetails.Match(responseSite2);
                objetoRetorno.Name = Name.Match(matchesMatrixDetails.Groups[0].Value).Groups[1].Value;
                objetoRetorno.MatrixID = MatrixID.Match(matchesMatrixDetails.Groups[0].Value).Groups[1].Value;
                objetoRetorno.Class = Class.Match(matchesMatrixDetails.Groups[0].Value).Groups[1].Value;
                objetoRetorno.Family = Family.Match(matchesMatrixDetails.Groups[0].Value).Groups[1].Value;
                objetoRetorno.Collection = Collection.Match(matchesMatrixDetails.Groups[0].Value).Groups[2].Value;
                objetoRetorno.Taxon = Taxon.Match(matchesMatrixDetails.Groups[0].Value).Groups[1].Value;
                objetoRetorno.Species = Species.Match(matchesMatrixDetails.Groups[0].Value).Groups[3].Value;
                objetoRetorno.DataType = DataType.Match(matchesMatrixDetails.Groups[0].Value).Groups[1].Value;
                objetoRetorno.Validation = Validation.Match(matchesMatrixDetails.Groups[0].Value).Groups[1].Value;
                objetoRetorno.UniprotID = UniprotID.Match(matchesMatrixDetails.Groups[0].Value).Groups[1].Value;
                objetoRetorno.Link = link;

                listaRetorno.Add(objetoRetorno);
            }

            return listaRetorno;
        }

        private static void GravarDadosExcel(List<ObjetoMeme> ListaObjeto, string outputFile)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Result");

                var pagina1WS = excel.Workbook.Worksheets[0];
                pagina1WS.Cells[$"A1"].Value = "Link";
                pagina1WS.Cells[$"A1"].Style.Font.Bold = true;
                pagina1WS.Cells[$"B1"].Value = "Name";
                pagina1WS.Cells[$"B1"].Style.Font.Bold = true;
                pagina1WS.Cells[$"C1"].Value = "MatrixID";
                pagina1WS.Cells[$"C1"].Style.Font.Bold = true;
                pagina1WS.Cells[$"D1"].Value = "Class";
                pagina1WS.Cells[$"D1"].Style.Font.Bold = true;
                pagina1WS.Cells[$"E1"].Value = "Family";
                pagina1WS.Cells[$"E1"].Style.Font.Bold = true;
                pagina1WS.Cells[$"F1"].Value = "Collection";
                pagina1WS.Cells[$"F1"].Style.Font.Bold = true;
                pagina1WS.Cells[$"G1"].Value = "Taxon";
                pagina1WS.Cells[$"G1"].Style.Font.Bold = true;
                pagina1WS.Cells[$"H1"].Value = "Species";
                pagina1WS.Cells[$"H1"].Style.Font.Bold = true;
                pagina1WS.Cells[$"I1"].Value = "DataType";
                pagina1WS.Cells[$"I1"].Style.Font.Bold = true;
                pagina1WS.Cells[$"J1"].Value = "Validation";
                pagina1WS.Cells[$"J1"].Style.Font.Bold = true;
                pagina1WS.Cells[$"K1"].Value = "UniprotID";
                pagina1WS.Cells[$"K1"].Style.Font.Bold = true;

                int pLinhaCorrespondente = 2;
                foreach (ObjetoMeme pObjeto in ListaObjeto)
                {
                    pagina1WS.Cells[$"A{pLinhaCorrespondente}"].Value = pObjeto.Link;
                    pagina1WS.Cells[$"B{pLinhaCorrespondente}"].Value = pObjeto.Name;
                    pagina1WS.Cells[$"C{pLinhaCorrespondente}"].Value = pObjeto.MatrixID;
                    pagina1WS.Cells[$"D{pLinhaCorrespondente}"].Value = pObjeto.Class;
                    pagina1WS.Cells[$"E{pLinhaCorrespondente}"].Value = pObjeto.Family;
                    pagina1WS.Cells[$"F{pLinhaCorrespondente}"].Value = pObjeto.Collection;
                    pagina1WS.Cells[$"G{pLinhaCorrespondente}"].Value = pObjeto.Taxon;
                    pagina1WS.Cells[$"H{pLinhaCorrespondente}"].Value = pObjeto.Species;
                    pagina1WS.Cells[$"I{pLinhaCorrespondente}"].Value = pObjeto.DataType;
                    pagina1WS.Cells[$"J{pLinhaCorrespondente}"].Value = pObjeto.Validation;
                    pagina1WS.Cells[$"K{pLinhaCorrespondente}"].Value = pObjeto.UniprotID;

                    pLinhaCorrespondente++;
                }
                excel.SaveAs(outputFile);
            }

        }
    }

    public class ObjetoMeme
    {
        public string Link { get; set; }
        public string Name { get; set; }
        public string MatrixID { get; set; }
        public string Class { get; set; }
        public string Family { get; set; }
        public string Collection { get; set; }
        public string Taxon { get; set; }
        public string Species { get; set; }
        public string DataType { get; set; }
        public string Validation { get; set; }
        public string UniprotID { get; set; }
    }
}