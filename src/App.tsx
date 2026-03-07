import { useEffect, useMemo, useState } from "react"
import { LockKeyhole, LogIn, ShieldCheck, ChartNoAxesColumn } from "lucide-react"

import { HeroGeometric } from "@/components/ui/shape-landing-hero"

export default function App() {
  const [username, setUsername] = useState("")
  const [password, setPassword] = useState("")
  const [error, setError] = useState("")
  const [isSubmitting, setIsSubmitting] = useState(false)

  const destination = useMemo(() => "/workspace/index.html", [])

  useEffect(() => {
    let cancelled = false

    async function checkSession() {
      try {
        const response = await fetch("/api/session", {
          credentials: "same-origin",
        })

        if (!cancelled && response.ok) {
          window.location.replace(destination)
        }
      } catch {
        // Local dev may not have the Vercel session endpoint.
      }
    }

    void checkSession()

    return () => {
      cancelled = true
    }
  }, [destination])

  const handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    setError("")

    const normalizedUsername = username.trim().toLowerCase()

    if (!normalizedUsername || !password) {
      setError("Enter both username and password to continue.")
      return
    }

    setIsSubmitting(true)

    try {
      const response = await fetch("/api/login", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        credentials: "same-origin",
        body: JSON.stringify({
          username: normalizedUsername,
          password,
        }),
      })

      const payload = await response.json().catch(() => ({}))

      if (!response.ok) {
        setError(typeof payload.error === "string" ? payload.error : "Login failed.")
        setPassword("")
        return
      }

      window.location.assign(destination)
    } catch {
      setError("Login is unavailable right now. Try again in production or configure the Vercel API locally.")
    } finally {
      setIsSubmitting(false)
    }
  }

  return (
    <main className="relative min-h-screen overflow-hidden bg-neutral-950 text-white">
      <div className="absolute inset-0">
        <HeroGeometric showContent={false} />
      </div>

      <div className="relative z-10 flex min-h-screen items-center justify-center px-6 py-10">
        <div className="grid w-full max-w-6xl gap-8 rounded-[2rem] border border-white/10 bg-black/45 p-4 shadow-[0_40px_120px_rgba(0,0,0,0.55)] backdrop-blur-xl lg:grid-cols-[1.1fr_0.9fr] lg:p-6">
          <section className="flex flex-col justify-between rounded-[1.6rem] border border-white/8 bg-white/6 p-8 lg:p-10">
            <div className="space-y-6">
              <div className="inline-flex items-center gap-2 rounded-full border border-orange-400/20 bg-orange-500/10 px-4 py-2 text-xs font-semibold uppercase tracking-[0.35em] text-orange-200">
                <ShieldCheck className="h-4 w-4" />
                Earthsoft Control Layer
              </div>

              <div className="space-y-4">
                <h1 className="max-w-xl text-4xl font-semibold tracking-[-0.05em] text-white md:text-6xl">
                  Project workspace access for live earthwork planning.
                </h1>
                <p className="max-w-2xl text-base leading-7 text-white/68 md:text-lg">
                  Secure the operational dashboard, keep project calculations behind a protected session, and route
                  authorized users directly into the workspace.
                </p>
              </div>
            </div>

            <div className="mt-8 grid gap-4 sm:grid-cols-3">
              <FeatureCard
                icon={<LockKeyhole className="h-5 w-5" />}
                title="Protected entry"
                copy="Session-based access gate for the project dashboard."
              />
              <FeatureCard
                icon={<ChartNoAxesColumn className="h-5 w-5" />}
                title="Live controls"
                copy="Open the earthwork calculator, diagrams, and project exports after sign-in."
              />
              <FeatureCard
                icon={<ShieldCheck className="h-5 w-5" />}
                title="Deployment-ready"
                copy="Built for Vercel-hosted access with a server-side password check."
              />
            </div>
          </section>

          <div className="rounded-[1.6rem] border border-white/10 bg-neutral-950/80 p-6 shadow-inner shadow-black/30 lg:p-8">
            <section className="mx-auto flex h-full max-w-md flex-col justify-center">
              <div className="mb-8 space-y-3">
                <p className="text-sm uppercase tracking-[0.35em] text-white/45">Restricted Login</p>
                <h2 className="text-3xl font-semibold tracking-[-0.04em]">Workspace access</h2>
                <p className="text-sm leading-6 text-white/60">
                  Use your authorized project login to open the Earthsoft dashboard.
                </p>
              </div>

              <form className="space-y-5" onSubmit={handleSubmit}>
                <div className="space-y-4">
                  <label className="block space-y-2">
                    <span className="text-sm font-medium text-white/70">Username</span>
                    <input
                      className="w-full rounded-2xl border border-white/10 bg-white/5 px-4 py-3 text-white outline-none transition focus:border-orange-400/60 focus:bg-white/10"
                      value={username}
                      onChange={(event) => setUsername(event.target.value)}
                      autoComplete="username"
                      placeholder="Enter your username"
                    />
                  </label>

                  <label className="block space-y-2">
                    <span className="text-sm font-medium text-white/70">Password</span>
                    <input
                      className="w-full rounded-2xl border border-white/10 bg-white/5 px-4 py-3 text-white outline-none transition focus:border-orange-400/60 focus:bg-white/10"
                      value={password}
                      onChange={(event) => setPassword(event.target.value)}
                      autoComplete="current-password"
                      placeholder="Enter your password"
                      type="password"
                    />
                  </label>
                </div>

                <div className="min-h-6 text-sm text-red-300">{error}</div>

                <button
                  type="submit"
                  disabled={isSubmitting}
                  className="inline-flex w-full items-center justify-center gap-2 rounded-2xl bg-white px-4 py-3 font-semibold text-black transition hover:bg-orange-100 disabled:cursor-not-allowed disabled:opacity-70"
                >
                  <LogIn className="h-4 w-4" />
                  {isSubmitting ? "Opening Workspace..." : "Enter Earthsoft"}
                </button>
              </form>
            </section>
          </div>
        </div>
      </div>
    </main>
  )
}

type FeatureCardProps = {
  icon: React.ReactNode
  title: string
  copy: string
}

function FeatureCard({ icon, title, copy }: FeatureCardProps) {
  return (
    <article className="rounded-3xl border border-white/10 bg-black/20 p-5">
      <div className="mb-4 inline-flex h-10 w-10 items-center justify-center rounded-2xl bg-orange-500/15 text-orange-200">
        {icon}
      </div>
      <h3 className="text-base font-semibold text-white">{title}</h3>
      <p className="mt-2 text-sm leading-6 text-white/60">{copy}</p>
    </article>
  )
}
