import { signIn } from "next-auth/react";
import { useRouter } from "next/router";

export default function Login() {
  const router = useRouter();
  const error = router.query.error;
  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#0e1a2b", color: "#fff", fontFamily: "'Barlow', sans-serif" }}>
      <div style={{ textAlign: "center", maxWidth: 420, padding: 24 }}>
        <div style={{ fontSize: 12, letterSpacing: 3, color: "#c5a662", marginBottom: 6 }}>TELETHERAPEUTICS HEALTH INC.</div>
        <h1 style={{ fontFamily: "'Barlow Condensed', sans-serif", fontWeight: 800, fontSize: 56, letterSpacing: 1, margin: 0 }}>
          <span style={{ color: "#fff" }}>TEAM</span><span style={{ color: "#c5a662" }}>LINK</span>
        </h1>
        <p style={{ color: "#a3aab5", margin: "12px 0 28px", lineHeight: 1.5 }}>Sign in with your @teletherapeuticsinc.com account to continue.</p>
        <button onClick={() => signIn("google", { callbackUrl: "/" })} style={{ background: "#c5a662", color: "#0e1a2b", border: "none", padding: "14px 36px", fontSize: 15, fontWeight: 700, letterSpacing: 1, borderRadius: 6, cursor: "pointer" }}>
          CONTINUE WITH GOOGLE
        </button>
        {error ? (
          <p style={{ color: "#ef9a9a", marginTop: 20, fontSize: 13 }}>
            {error === "AccessDenied" ? "Your account isn't allowed. Use your @teletherapeuticsinc.com email." : "Sign-in failed. Try again."}
          </p>
        ) : null}
      </div>
    </div>
  );
}
