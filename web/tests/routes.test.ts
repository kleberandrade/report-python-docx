import { afterEach, describe, expect, it, vi } from "vitest";
const storage = vi.hoisted(() => ({
  saveResponse: vi.fn(),
  readResponses: vi.fn(),
}));
vi.mock("../lib/database", async (original) => ({
  ...(await original<typeof import("../lib/database")>()),
  ...storage,
}));
import { POST as submit } from "../app/api/responses/route";
import { POST as download } from "../app/api/export/route";
import { authorized } from "../lib/database";
const item = {
  id: "e550419b-d01c-432e-91c8-8a605d44a998",
  stage: "alpha",
  game: "Órbita",
  scores: [8, 6, 10],
  play: "Sim",
  pay: "Menos R$ 10,00",
  comments: ["Ótima ação"],
};
const key = "test-only-administrative-key-123456";
function request(body: unknown, token?: string, origin = "http://localhost") {
  return new Request("http://localhost/api/responses", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      origin,
      ...(token ? { authorization: `Bearer ${token}` } : {}),
    },
    body: JSON.stringify(body),
  });
}
afterEach(() => {
  vi.unstubAllEnvs();
  vi.clearAllMocks();
});
describe("submission endpoint", () => {
  it("refuses cross-origin and invalid inputs before saving", async () => {
    expect(
      (await submit(request(item, undefined, "https://example.org"))).status,
    ).toBe(403);
    expect(
      (await submit(request({ ...item, scores: [-1, 6, 10] }))).status,
    ).toBe(400);
    expect(storage.saveResponse).not.toHaveBeenCalled();
  });
  it("refuses oversized streamed bodies", async () => {
    expect(
      (await submit(request({ ...item, game: "x".repeat(16000) }))).status,
    ).toBe(400);
    expect(storage.saveResponse).not.toHaveBeenCalled();
  });
  it("never reports success without storage", async () => {
    vi.stubEnv("DATABASE_URL", "");
    expect((await submit(request(item))).status).toBe(503);
    expect(storage.saveResponse).not.toHaveBeenCalled();
  });
  it("confirms only after saving", async () => {
    vi.stubEnv("DATABASE_URL", "postgres://test");
    storage.saveResponse.mockResolvedValueOnce(undefined);
    const response = await submit(request(item));
    expect(response.status).toBe(201);
    expect(await response.json()).toEqual({ saved: true });
    expect(storage.saveResponse).toHaveBeenCalledWith(item);
  });
  it("handles database errors without leaking credentials", async () => {
    vi.stubEnv("DATABASE_URL", "postgres://test");
    storage.saveResponse.mockRejectedValueOnce(
      new Error("private database detail"),
    );
    const response = await submit(request(item));
    expect(response.status).toBe(503);
    expect(await response.text()).not.toContain("private database");
  });
});
describe("protected export", () => {
  it("requires a configured strong key and exact comparison", () => {
    vi.stubEnv("ADMIN_PASSWORD", "short");
    expect(authorized("short")).toBe(false);
    vi.stubEnv("ADMIN_PASSWORD", key);
    expect(authorized(key)).toBe(true);
    expect(authorized("x".repeat(key.length))).toBe(false);
    expect(authorized(null)).toBe(false);
  });
  it("never queries answers for anonymous users", async () => {
    expect((await download(request({ stage: "alpha" }))).status).toBe(401);
    expect(storage.readResponses).not.toHaveBeenCalled();
  });
  it("handles empty results and invalid stages", async () => {
    vi.stubEnv("ADMIN_PASSWORD", key);
    storage.readResponses.mockResolvedValueOnce([]);
    expect((await download(request({ stage: "alpha" }, key))).status).toBe(404);
    expect((await download(request({ stage: "final" }, key))).status).toBe(400);
  });
  it("exports only requested stage with no-store and a safe filename", async () => {
    vi.stubEnv("ADMIN_PASSWORD", key);
    storage.readResponses.mockResolvedValueOnce([item]);
    const response = await download(request({ stage: "alpha" }, key));
    expect(response.status).toBe(200);
    expect(response.headers.get("cache-control")).toBe("no-store");
    expect(response.headers.get("content-disposition")).toContain(
      "alpha-evaluation.csv",
    );
    expect(await response.text()).toContain("Órbita,8,6,10");
    expect(storage.readResponses).toHaveBeenCalledWith("alpha");
  });
});

describe("reverse proxy origin handling", () => {
  it("matches the browser Host header when Next normalizes the internal URL", async () => {
    vi.stubEnv("DATABASE_URL", "");
    const incoming = new Request("http://localhost:8793/api/responses", {
      method: "POST",
      headers: {
        host: "127.0.0.1:8793",
        origin: "http://127.0.0.1:8793",
        "content-type": "application/json",
      },
      body: JSON.stringify(item),
    });
    expect((await submit(incoming)).status).toBe(503);
  });
});
