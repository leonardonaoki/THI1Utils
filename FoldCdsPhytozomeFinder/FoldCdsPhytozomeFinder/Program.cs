using OfficeOpenXml;
using PuppeteerSharp;
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace FoldCdsPhytozomeFinder
{
    public class Program
    {
        static void Main(string[] args)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var excel = new ExcelPackage())
                {
                    //Here was utilized an excel with the first column as 'ID Database', the second as 'Database (Phytozome)' and the third as ' ID Gene'
                    using (var stream = File.OpenRead(args[0]))
                        excel.Load(stream);

                    var pagina1WS = excel.Workbook.Worksheets[0];

                    int contadorCell = 2;
                    foreach (var cell in pagina1WS.Cells["A2:A28"])
                    {
                        string IDGENE = pagina1WS.Cells[$"C{contadorCell}"].Value.ToString();

                        sb.AppendLine(">");
                        string informacoesPhytozome = baterSite(cell.Hyperlink.AbsoluteUri).Result;
                        if (!string.IsNullOrWhiteSpace(informacoesPhytozome))
                            sb.Append(informacoesPhytozome);
                        else
                            Console.WriteLine($"Não foi possível carregar as informações do link contido na celula:{contadorCell}");

                        sb.AppendLine(Environment.NewLine);

                        contadorCell++;

                        int index = informacoesPhytozome.IndexOf("\n");
                        string informacaoFinal = informacoesPhytozome.Substring(index + 1);
                        string bicoPato = Environment.NewLine + $">{IDGENE}" + Environment.NewLine;
                        string final = bicoPato + informacaoFinal;

                        File.AppendAllText(args[1], final);
                    }

                    Console.WriteLine("Finalizado");
                    Console.ReadLine();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(@"Por favor preencha corretamente os parâmetros separados por espaços. 
                Lembrando, 1º Parametro = Caminho do Excel em que a 1º coluna = 'ID Database' , 2º coluna = 'Database (Phytozome) e 3º coluna = 'ID Gene''
                2º Parametro = Caminho do Arquivo de saída, exemplo 'C:\Users\Teste\Desktop\Resultado.txt' ");
                Console.WriteLine("Erro Original: " + e);
            }

        }

        /// <summary>
        /// Método utilizado para obter o cds do site do phytozome, é possível substituir o 'fold-cds' para obter o resultado desejado
        /// </summary>
        /// <param name="pUrl">Url do phytozone</param>
        /// <returns></returns>
        public static async Task<string> baterSite(string pUrl)
        {
            await new BrowserFetcher().DownloadAsync(BrowserFetcher.DefaultRevision);
            using (var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = false }))
            using (var page = await browser.NewPageAsync())
            {
                await page.SetViewportAsync(new ViewPortOptions() { Width = 1280, Height = 600 });
                await page.GoToAsync(pUrl);
                await page.WaitForTimeoutAsync(5000);
                string header = await page.EvaluateExpressionAsync<string>("document.getElementById('fold-cds').nextElementSibling.nextElementSibling.children[0].innerText");

                return header;
            }

        }
    }
}

