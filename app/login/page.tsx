"use client";

import { useState } from "react";
import { useRouter } from "next/navigation";
import { login } from "../actions";

export default function LoginPage() {
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const router = useRouter();

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    setIsLoading(true);
    setError("");

    try {
      const success = await login(password);
      if (success) {
        router.push("/");
      } else {
        setError("Incorrect password");
      }
    } catch (err) {
      setError("An error occurred. Please try again.");
    } finally {
      setIsLoading(false);
    }
  }

  return (
    <main className="page-shell" style={{ display: "flex", alignItems: "center", justifyContent: "center", minHeight: "100vh" }}>
      <div className="panel" style={{ maxWidth: 400, width: "100%", padding: 40 }}>
        <section className="hero" style={{ marginBottom: 32, padding: 0 }}>
          <p className="eyebrow">Secure Access</p>
          <h1 style={{ fontSize: "1.5rem", marginBottom: 8 }}>Patient XML Reader</h1>
          <p style={{ fontSize: "0.9rem", color: "var(--fg-muted)" }}>
            Please enter your access password to view patient records.
          </p>
        </section>

        <form onSubmit={handleSubmit} className="stack">
          <div>
            <label className="label" htmlFor="password">Password</label>
            <input
              id="password"
              className="search-input"
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="••••••••"
              required
              autoFocus
            />
          </div>

          {error && (
            <div className="warning" style={{ fontSize: "0.85rem", padding: "8px 12px" }}>
              {error}
            </div>
          )}

          <button 
            className="button" 
            type="submit" 
            disabled={isLoading || !password}
            style={{ width: "100%", justifyContent: "center", marginTop: 8 }}
          >
            {isLoading ? "Verifying..." : "Access Reader"}
          </button>
        </form>

        <p style={{ marginTop: 32, fontSize: "0.75rem", color: "var(--fg-muted)", textAlign: "center" }}>
          Protected health information. For authorized use only.
        </p>
      </div>
    </main>
  );
}
