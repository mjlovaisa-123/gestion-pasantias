import React, { useEffect, useState } from "react";
import { supabase } from "./lib/supabaseClient";
import App from "./App.jsx";

const fontImport = `
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Serif:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');
`;

function LoginScreen({ onLogin }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    setLoading(false);
    if (error) {
      setError("Email o contraseña incorrectos.");
      return;
    }
    onLogin();
  };

  return (
    <div style={{ minHeight: "100vh", background: "#F6F7F5", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'IBM Plex Sans', sans-serif", padding: 20 }}>
      <style>{fontImport}</style>
      <form onSubmit={submit} style={{ background: "#fff", borderRadius: 12, padding: 32, width: "100%", maxWidth: 360, boxShadow: "0 20px 60px rgba(0,0,0,0.12)" }}>
        <div style={{ fontFamily: "'IBM Plex Serif', serif", fontWeight: 700, fontSize: 20, color: "#1B2A4A", marginBottom: 4 }}>
          Gestión de Pasantías
        </div>
        <div style={{ fontSize: 13, color: "#5B6158", marginBottom: 24 }}>Iniciá sesión para continuar</div>

        <label style={{ fontSize: 12, fontWeight: 600, color: "#5B6158", marginBottom: 4, display: "block" }}>Email</label>
        <input
          type="email"
          required
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          style={{ width: "100%", padding: "10px 12px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 14, boxSizing: "border-box", marginBottom: 14 }}
        />

        <label style={{ fontSize: 12, fontWeight: 600, color: "#5B6158", marginBottom: 4, display: "block" }}>Contraseña</label>
        <input
          type="password"
          required
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          style={{ width: "100%", padding: "10px 12px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 14, boxSizing: "border-box", marginBottom: error ? 8 : 20 }}
        />

        {error && <div style={{ fontSize: 12.5, color: "#A6432D", marginBottom: 14 }}>{error}</div>}

        <button
          type="submit"
          disabled={loading}
          style={{ width: "100%", padding: "11px", borderRadius: 6, border: "none", background: "#1B2A4A", color: "#fff", fontSize: 14, fontWeight: 600, cursor: loading ? "default" : "pointer", opacity: loading ? 0.7 : 1 }}
        >
          {loading ? "Ingresando…" : "Ingresar"}
        </button>

        <div style={{ fontSize: 11.5, color: "#8A9088", marginTop: 16, textAlign: "center" }}>
          Si no tenés usuario, pedile al administrador del sistema que te lo cree.
        </div>
      </form>
    </div>
  );
}

export default function AuthGate() {
  const [session, setSession] = useState(undefined); // undefined = cargando, null = sin sesión

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => listener.subscription.unsubscribe();
  }, []);

  if (session === undefined) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "sans-serif", color: "#5B6158" }}>
        Cargando…
      </div>
    );
  }

  if (!session) {
    return <LoginScreen onLogin={() => {}} />;
  }

  return <App onLogout={() => supabase.auth.signOut()} userEmail={session.user?.email} />;
}
