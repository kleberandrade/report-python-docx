import Link from "next/link";
export default function NotFound() {
  return (
    <section className="wrap empty">
      <span className="eyebrow">PÁGINA NÃO ENCONTRADA</span>
      <h1>Vamos voltar ao jogo?</h1>
      <p>Esta avaliação não existe.</p>
      <Link className="button" href="/">
        Ir para o início
      </Link>
    </section>
  );
}
