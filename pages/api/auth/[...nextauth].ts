import NextAuth from "next-auth";
import GoogleProvider from "next-auth/providers/google";

const allowedDomains = (process.env.ALLOWED_GOOGLE_DOMAINS || "")
  .split(",")
  .map((d) => d.trim().toLowerCase())
  .filter(Boolean);

export default NextAuth({
  providers: [
    GoogleProvider({
      clientId: process.env.GOOGLE_CLIENT_ID || "",
      clientSecret: process.env.GOOGLE_CLIENT_SECRET || "",
    }),
  ],
  callbacks: {
    async signIn({ user }) {
      if (allowedDomains.length === 0) return true;
      const email = (user.email || "").toLowerCase();
      const domain = email.split("@")[1];
      return Boolean(domain && allowedDomains.includes(domain));
    },
  },
  pages: {
    signIn: "/login",
    error: "/login",
  },
});
