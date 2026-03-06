import { useMemo, useState } from "react"
import { LockKeyhole, LogIn, ShieldCheck, ChartNoAxesColumn } from "lucide-react"

import { HeroGeometric } from "@/components/ui/shape-landing-hero"

const AUTH_STORAGE_KEY = "earthsoft_auth_session"
const AUTH_USERNAME = "admin"
const AUTH_PASSWORD = "earthsoft123"

export default function App() {
  const [username, setUsername] = useState(AUTH_USERNAME)
  const [password, setPassword] = useState(AUTH_PASSWORD)
  const [error, setError] = useState("")
  const [isSubmitting, setIsSubmitting] = useState(false)

  const destination = useMemo(() => "/workspace/index.html", [])

  const handleSubmit = (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    setError("")

    const normalizedUsername = username.trim().toLowerCase()
    const normalizedPassword = password.trim()

    if (!normalizedUsername || !normalizedPassword) {
      setError("Enter both username and password to continue.")
      return
    }

    if (normalizedUsername !== AUTH_USERNAME || normalizedPassword !== AUTH_PASSWORD) {
      setError("Invalid username or password.")
      setPassword("")
      return
    }

    setIsSubmitting(true)
    sessionStorage.setItem(
      AUTH_STORAGE_KEY,
      JSON.stringify({
        authenticated: true,
        user: AUTH_USERNAME,
      }),
    )
    window.location.assign(destination)
  }

  return (
    <main className="relative min-h-screen overflow-hidden bg-black text-white">
      <div className="absolute inset-0">
        <HeroGeometric
          badge="Earthwork Suite"
          title1="Railway Earthwork"
          title2="Planning Workspace"
        />
      </div>

      <div className="relative z-10 flex min-h-screen items-center justify-center p-6">
        <div className="w-full max-w-3xl rounded-[2rem] border border-white/10 bg-black/45 p-6 shadow-2xl shadow-black/50 backdrop-blur-2xl sm:p-8 lg:p-10">
          <div className="grid gap-8 lg:grid-cols-[0.8fr_1.2fr] lg:items-start">
            <section className="space-y-6">
              <div className="flex items-center gap-5">
                <div className="inline-flex h-16 w-16 shrink-0 items-center justify-center overflow-hidden rounded-2xl border border-white/10 bg-white/6 shadow-2xl shadow-black/30 backdrop-blur-xl">
                  <img
                    src="/workspace/assets/logo_silhouette.png"
                    alt="Earthsoft logo"
                    className="h-full w-full object-cover"
                  />
                </div>
                <div className="flex flex-col justify-center">
                  <span className="text-3xl font-bold leading-none tracking-[-0.06em] text-white sm:text-4xl">
                    Earthsoft
                  </span>
                  <span className="mt-1 text-xs font-medium tracking-[0.02em] text-white/68 sm:text-sm">
                    The Railway earthwork workspace.
                  </span>
                </div>
              </div>
              <div className="space-y-4">
                <p className="max-w-2xl text-sm leading-6 text-white/68 sm:text-base">
                  Review quantities, verify project data, and continue estimate preparation.
                </p>
              </div>
              <div className="flex flex-wrap gap-3">
                <div className="inline-flex items-center gap-2 rounded-full border border-white/10 bg-white/3 px-3 py-2 text-white/78 backdrop-blur-md">
                  <ShieldCheck className="h-3.5 w-3.5 text-white/70" />
                  <span className="text-[10px] sm:text-xs">Session access</span>
                </div>
                <div className="inline-flex items-center gap-3 rounded-full border border-white/10 bg-white/3 px-3 py-2 text-white/78 backdrop-blur-md">
                  <ChartNoAxesColumn className="h-3.5 w-3.5 text-white/70" />
                  <span className="text-[10px] sm:text-xs">Railway suite</span>
                </div>
              </div>
            </section>

            <section className="rounded-[1.75rem] border border-white/10 bg-black/55 p-6 shadow-2xl shadow-black/40 backdrop-blur-2xl sm:p-8">
              <form className="space-y-6" onSubmit={handleSubmit}>
                <div className="space-y-2">
                  <div className="inline-flex items-center gap-2 rounded-full border border-white/10 bg-white/5 px-3 py-1 text-xs uppercase tracking-[0.24em] text-white/55">
                    <LockKeyhole className="h-3.5 w-3.5" />
                    Secure Login
                  </div>
                  <h2 className="text-3xl font-semibold tracking-[-0.04em]">Workspace access</h2>
                  <p className="text-sm leading-6 text-white/60">
                    Use your authorized project login to open the Earthsoft dashboard.
                  </p>
                </div>

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
