import type { Metadata } from "next";
import Link from "next/link";
import "./globals.css";
export const metadata: Metadata = {
  title: {
    default: "Playtest · Avaliações de jogos",
    template: "%s · Playtest",
  },
  description:
    "Questionários acadêmicos alfa e beta para ouvir quem joga e transformar respostas em relatórios.",
  robots: { index: false, follow: false },
};
export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="pt-BR">
      <body>
        <a className="skip" href="#content">
          Pular para o conteúdo
        </a>
        <header className="site-header">
          <Link className="brand" href="/">
            <span className="brand-mark" aria-hidden="true">
              p.
            </span>
            playtest<span className="brand-note">caderno de avaliações</span>
          </Link>
          <nav aria-label="Navegação principal">
            <Link href="/relatorios">Ver relatórios</Link>
            <Link href="/resultados">
              Área docente <span aria-hidden="true">↗</span>
            </Link>
          </nav>
        </header>
        <main id="content">{children}</main>
        <footer className="site-footer">
          <span>Uma escuta melhor. Uma próxima versão melhor.</span>
          <span>Projeto de Kleber de Oliveira Andrade · MIT</span>
        </footer>
      </body>
    </html>
  );
}
