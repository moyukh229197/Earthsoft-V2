import { useEffect, useMemo, useState } from "react"
import { ArrowRightToLine, ChartNoAxesColumn, LockKeyhole, ShieldCheck } from "lucide-react"

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
        // Local dev may not expose the Vercel session endpoint.
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

      <div className="relative z-10 flex min-h-screen items-center justify-center px-5 py-8 md:px-8">
        <div className="relative w-full max-w-5xl overflow-hidden rounded-[1.8rem] border border-cyan-400/12 bg-[linear-gradient(135deg,rgba(40,90,255,0.08),rgba(3,8,24,0.03)_36%,rgba(8,16,34,0.03)_64%,rgba(32,196,163,0.08))] p-3 shadow-[0_24px_80px_rgba(0,0,0,0.42)] backdrop-blur-md lg:p-5">
          <div className="pointer-events-none absolute inset-0 rounded-[2rem] ring-1 ring-inset ring-white/6" />
          <div className="grid gap-6 rounded-[1.35rem] bg-black/8 p-2 backdrop-blur-sm lg:grid-cols-[1.15fr_0.75fr] lg:p-3">
            <section className="flex min-h-[420px] flex-col justify-between p-5 md:p-6">
              <div className="space-y-6">
                <div className="flex items-center gap-4">
                  <div className="h-20 w-20 overflow-hidden rounded-[1.45rem] border border-cyan-300/20 bg-[linear-gradient(180deg,rgba(32,179,255,0.28),rgba(68,84,255,0.18))] shadow-[0_10px_40px_rgba(43,120,255,0.25)]">
                    <img src="/workspace/assets/logo_silhouette.png" alt="Earthsoft" className="h-full w-full object-cover" />
                  </div>

                  <div className="space-y-2">
                    <p className="text-xs font-semibold uppercase tracking-[0.38em] text-cyan-200/90">Earthwork Suite</p>
                    <h1 className="text-3xl font-semibold tracking-[-0.06em] text-white md:text-5xl">Earthsoft Access</h1>
                  </div>
                </div>

                <p className="max-w-xl text-base leading-7 text-white/70">
                  Estimate quantities, review profiles, manage bridges and curves, and export project-ready outputs.
                </p>
              </div>

              <div className="grid gap-3 pt-6 sm:grid-cols-3">
                <MetaPill icon={<LockKeyhole className="h-4 w-4" />} text="Quantity calculations" />
                <MetaPill icon={<ChartNoAxesColumn className="h-4 w-4" />} text="Profiles and diagrams" />
                <MetaPill icon={<ShieldCheck className="h-4 w-4" />} text="Project-ready exports" />
              </div>
            </section>

            <section className="flex p-2 md:p-3">
              <form className="flex w-full flex-col justify-between rounded-[1.15rem] border border-cyan-300/10 bg-[linear-gradient(180deg,rgba(7,19,45,0.22),rgba(7,19,45,0.08))] p-4 backdrop-blur-md md:p-5" onSubmit={handleSubmit}>
                <div className="space-y-4">
                  <div className="space-y-2">
                    <h2 className="text-xl font-semibold tracking-[-0.04em] text-white md:text-2xl">Sign in</h2>
                    <p className="text-sm leading-6 text-white/58">Continue to your Earthsoft workspace.</p>
                  </div>

                  <label className="block space-y-2">
                    <span className="text-sm font-medium text-white/74">Username</span>
                    <input
                      className="w-full rounded-[1rem] border border-cyan-300/20 bg-[rgba(6,18,42,0.32)] px-4 py-3.5 text-white outline-none transition placeholder:text-white/30 focus:border-cyan-300/60 focus:bg-[rgba(8,28,62,0.42)]"
                      value={username}
                      onChange={(event) => setUsername(event.target.value)}
                      autoComplete="username"
                      placeholder="Enter your username"
                    />
                  </label>

                  <label className="block space-y-2">
                    <span className="text-sm font-medium text-white/74">Password</span>
                    <input
                      className="w-full rounded-[1rem] border border-white/10 bg-[rgba(6,18,42,0.26)] px-4 py-3.5 text-white outline-none transition placeholder:text-white/30 focus:border-cyan-300/50 focus:bg-[rgba(8,28,62,0.38)]"
                      value={password}
                      onChange={(event) => setPassword(event.target.value)}
                      autoComplete="current-password"
                      placeholder="Enter your password"
                      type="password"
                    />
                  </label>

                  <div className="min-h-6 text-sm text-red-300">{error}</div>
                </div>

                <button
                  type="submit"
                  disabled={isSubmitting}
                  className="mt-6 inline-flex w-full items-center justify-center gap-3 rounded-[1rem] bg-[linear-gradient(90deg,#3254ff,#2c67ff)] px-4 py-3.5 font-semibold text-white shadow-[0_12px_30px_rgba(50,84,255,0.35)] transition hover:brightness-110 disabled:cursor-not-allowed disabled:opacity-70"
                >
                  <ArrowRightToLine className="h-4 w-4" />
                  {isSubmitting ? "Opening Workspace..." : "Sign In"}
                </button>
              </form>
            </section>
          </div>
        </div>
      </div>
    </main>
  )
}

type MetaPillProps = {
  icon: React.ReactNode
  text: string
}

function MetaPill({ icon, text }: MetaPillProps) {
  return (
    <div className="inline-flex items-center gap-3 rounded-full border border-white/10 bg-black/8 px-4 py-2.5 text-sm text-white/72 backdrop-blur-sm">
      <span className="text-cyan-200">{icon}</span>
      <span>{text}</span>
    </div>
  )
}
