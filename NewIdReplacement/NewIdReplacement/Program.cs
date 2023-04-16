using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NewIdReplacement
{
    public class Program
    {
        static void Main(string[] args)
        {
            try
            {
                new NovaIdIdentificador().gerarNovaIDReplace(args[0], args[1], args[2]);
            }
            catch (Exception e)
            {
                Console.WriteLine(@"Por favor preencha corretamente os parâmetros separados por espaços. 
                Lembrando, 1º Parametro = TXT de representando um arquivo de referência chave valor separados por espaço e quebra de linha, exemplo:
                F_Byssochlamys spectabilis	PVAR5_2261
                F_Capronia epimyces	A1O3_00193
                F_Cladophialophora bantiana	Z519_06524
        
                2º Parametro = Caminho do Arquivo fasta de substituição, exemplo: 'C:\Users\Teste\Desktop\Fungos_substituir.fasta
                3º Parametro = Caminho do Arquivo de saída, exemplo 'C:\Users\Teste\Desktop\Fungos_result.txt' ");
                Console.WriteLine("Erro Original: " + e);
            }

        }
    }

    public class NovaIdIdentificador
    {
        public string NovaIdCerto { get; set; }
        public string IdentificadorASubstituir { get; set; }
        public bool Preenchido { get; set; }

        public void gerarNovaIDReplace(string txtFungos, string fungosFasta, string outputTXT)
        {
            var linhasArquivo = File.ReadAllLines(txtFungos);

            List<NovaIdIdentificador> listaNovaId = new List<NovaIdIdentificador>();

            foreach (var linha in linhasArquivo)
            {
                var splitLinha = linha.Split('\t');
                NovaIdIdentificador identificador = new NovaIdIdentificador()
                {
                    NovaIdCerto = splitLinha[0],
                    IdentificadorASubstituir = splitLinha[1]
                };
                listaNovaId.Add(identificador);
            }

            var LinhasArquivoASubstituir = File.ReadAllLines(fungosFasta);

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < LinhasArquivoASubstituir.Count(); i++)
            {
                if (LinhasArquivoASubstituir[i].StartsWith(">"))
                {
                    var retorno = listaNovaId.FirstOrDefault(p => LinhasArquivoASubstituir[i].Contains(p.IdentificadorASubstituir));

                    if (retorno != null && LinhasArquivoASubstituir[i].Contains(retorno.IdentificadorASubstituir))
                    {
                        retorno.Preenchido = true;
                        sb.Append(Environment.NewLine + ">" + retorno.NovaIdCerto + Environment.NewLine);
                    }
                    else
                        i = i + 1;
                }
                else
                    sb.Append(LinhasArquivoASubstituir[i]);
            }

            var sobrou = listaNovaId.Where(y => y.Preenchido == false);

            sb.Append(Environment.NewLine);
            foreach (var item in sobrou)
            {
                sb.Append("Faltou: " + item.NovaIdCerto + " - " + item.IdentificadorASubstituir + Environment.NewLine);
            }

            File.WriteAllText(outputTXT, sb.ToString().Replace(">Não Existe", string.Empty));
        }
    }
}
