"use client";
import { useRef, useState } from "react";
import type { Stage } from "../lib/evaluations";
import { responsePayload } from "../lib/form";
export function useSubmission(stage: Stage) {
  const [status, setStatus] = useState<"idle" | "saving" | "saved">("idle");
  const [error, setError] = useState("");
  const id = useRef<string | null>(null);
  async function submit(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (status === "saving") return;
    setError("");
    setStatus("saving");
    const data = new FormData(event.currentTarget);
    id.current ??= crypto.randomUUID();
    const payload = responsePayload(data, stage, id.current);
    try {
      const response = await fetch("/api/responses", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const result = await response.json();
      if (!response.ok || !result.saved)
        throw new Error(
          result.error || "Não foi possível salvar. Tente novamente.",
        );
      setStatus("saved");
    } catch (failure) {
      setError(
        failure instanceof Error
          ? failure.message
          : "Conexão interrompida. Tente novamente.",
      );
      setStatus("idle");
    }
  }

  return {
    status,
    error,
    submit,
    reset: () => {
      id.current = null;
      setStatus("idle");
    },
  };
}
