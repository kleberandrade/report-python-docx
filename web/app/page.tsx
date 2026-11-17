import Link from "next/link";
export default function Home() {
  return (
    <>
      <section className="hero wrap">
        <div className="eyebrow">JOGOS DIGITAIS / AVALIAÇÃO ACADÊMICA</div>
        <h1>
          Todo jogo evolui
          <br />
          com quem <em>joga.</em>
        </h1>
        <div className="hero-bottom">
          <p>
            Da primeira ideia à experiência jogável.
            <br />
            Compartilhe sua percepção e ajude a equipe a escolher o próximo
            passo.
          </p>
          <span className="edition">
            DUAS ETAPAS
            <br />
            <strong>uma conversa contínua</strong>
          </span>
        </div>
      </section>
      <section className="stages wrap" aria-label="Escolha a avaliação">
        <Link className="stage-card alpha" href="/avaliar/alpha">
          <div className="card-top">
            <span className="pill">01 / CONCEITO</span>
            <span aria-hidden="true">↗</span>
          </div>
          <h2>Avaliação alfa</h2>
          <p>
            A ideia desperta interesse? Avalie originalidade, diversão e
            potencial de mercado.
          </p>
          <div className="card-bottom">
            <span>3 critérios · cerca de 3 min</span>
            <strong>Avaliar a ideia →</strong>
          </div>
        </Link>
        <Link className="stage-card beta" href="/avaliar/beta">
          <div className="card-top">
            <span className="pill">02 / EXPERIÊNCIA</span>
            <span aria-hidden="true">↗</span>
          </div>
          <h2>Avaliação beta</h2>
          <p>
            Como foi jogar? Avalie gráficos, áudio, interface, jogabilidade e
            história.
          </p>
          <div className="card-bottom">
            <span>5 critérios · cerca de 5 min</span>
            <strong>Avaliar o jogo →</strong>
          </div>
        </Link>
      </section>
      <section className="explanation wrap">
        <span className="eyebrow">DO PLAYTEST AO RELATÓRIO</span>
        <div>
          <h2>
            Respostas que viram
            <br />
            próximos passos.
          </h2>
          <p>
            A equipe docente reúne as avaliações de cada jogo em um documento
            com médias, distribuição de respostas e comentários completos.
          </p>
          <Link className="text-link" href="/relatorios">
            Confira os relatórios completos de exemplo →
          </Link>
        </div>
        <p className="privacy-note">
          Não pedimos nome nem email. Evite identificar pessoas nos comentários.
          Na etapa beta, idade e gênero ajudam a contextualizar a experiência;
          você pode optar por não informar o gênero.
        </p>
      </section>
    </>
  );
}
