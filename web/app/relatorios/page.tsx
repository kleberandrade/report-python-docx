import Image from "next/image";
export const metadata = { title: "Relatórios completos de exemplo" };
export default function Reports() {
  return (
    <div className="wrap reports-page">
      <div className="eyebrow">EXEMPLOS COMPLETOS / DADOS FICTÍCIOS</div>
      <h1>
        Da resposta
        <br />
        ao documento.
      </h1>
      <p className="intro">
        Confira todas as páginas dos exemplos alfa e beta. Cada documento tem
        uma página e usa duas avaliações fictícias do jogo Órbita.
      </p>
      <div className="report-grid">
        {["alpha", "beta"].map((stage) => (
          <section key={stage}>
            <div className="report-toolbar">
              <h2>Relatório {stage === "alpha" ? "alfa" : "beta"}</h2>
              <a href={`/examples/${stage}-report.docx`} download>
                Baixar DOCX ↓
              </a>
            </div>
            <a
              href={`/examples/${stage}-report.png`}
              target="_blank"
              rel="noreferrer"
              aria-label={`Abrir relatório ${stage} completo em tamanho original`}
            >
              <Image
                src={`/examples/${stage}-report.png`}
                alt={`Página completa do relatório ${stage} do jogo Órbita, com notas, distribuição das respostas e comentários fictícios.`}
                sizes="(max-width: 800px) 100vw, 50vw"
                width={1414}
                height={2000}
              />
            </a>
          </section>
        ))}
      </div>
    </div>
  );
}
