import { ExportForm } from "../../components/export-form";
export const metadata = { title: "Área docente" };
export default function Results() {
  return (
    <div className="form-wrap">
      <div className="form-header">
        <span className="eyebrow">ACESSO RESTRITO / EQUIPE DOCENTE</span>
        <h1>
          O próximo passo<span className="title-dot">.</span>
        </h1>
        <p>
          Exporte as respostas de uma etapa e gere um relatório DOCX para cada
          jogo com a ferramenta do projeto.
        </p>
      </div>
      <ExportForm />
      <section className="export-instructions">
        <h2>Do CSV ao relatório</h2>
        <p>
          Depois de baixar o arquivo, execute no computador com o gerador
          instalado:
        </p>
        <code>class-report alpha-evaluation.csv -o reports/alpha</code>
        <p>Para a etapa beta:</p>
        <code>
          class-report beta-evaluation.csv --stage beta -o reports/beta
        </code>
        <p>
          Os relatórios incluem todas as respostas exportadas. A exportação
          exige uma chave configurada pela pessoa responsável. Dados individuais
          não ficam disponíveis publicamente.
        </p>
      </section>
    </div>
  );
}
