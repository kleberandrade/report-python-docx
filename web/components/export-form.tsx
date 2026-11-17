"use client";
import { useState } from "react";
export function ExportForm() {
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("");
  async function download(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setBusy(true);
    setMessage("");
    const data = new FormData(event.currentTarget);
    const stage = data.get("stage");
    try {
      const response = await fetch("/api/export", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${data.get("key")}`,
        },
        body: JSON.stringify({ stage }),
      });
      if (!response.ok) {
        const result = await response.json();
        throw new Error(result.error);
      }
      const url = URL.createObjectURL(await response.blob());
      const a = document.createElement("a");
      a.href = url;
      a.download = `${stage}-evaluation.csv`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      setMessage(
        `${response.headers.get("X-Response-Count")} respostas exportadas. Nenhuma resposta foi removida do banco.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Não foi possível exportar.",
      );
    } finally {
      setBusy(false);
    }
  }
  return (
    <form className="export-panel" onSubmit={download}>
      <label htmlFor="stage">Etapa</label>
      <select id="stage" name="stage" defaultValue="alpha">
        <option value="alpha">Avaliação alfa</option>
        <option value="beta">Avaliação beta</option>
      </select>
      <label htmlFor="key">Chave da equipe docente</label>
      <input
        id="key"
        name="key"
        type="password"
        required
        autoComplete="off"
        minLength={24}
        aria-describedby="key-help"
      />
      <p className="field-help" id="key-help">
        A chave não é salva neste navegador. Solicite-a à pessoa responsável
        pela aplicação.
      </p>
      <button className="button" disabled={busy}>
        {busy ? "Preparando exportação…" : "Exportar respostas em CSV ↓"}
      </button>
      <p role="status" className="status-text">
        {message}
      </p>
    </form>
  );
}
