import { useMemo, useState } from "react"
import { ChartNoAxesColumn, LockKeyhole, LogIn, ShieldCheck } from "lucide-react"

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

      <div className="relative z-10 flex min-h-screen items-center justify-center p-4 sm:p-6">
        <div className="w-full max-w-3xl rounded-[1.75rem] border border-white/8 bg-[linear-gradient(135deg,rgba(8,24,66,0.48),rgba(2,10,28,0.34)_40%,rgba(2,8,20,0.4)_100%)] p-3 shadow-2xl shadow-black/45 backdrop-blur-xl sm:p-4 lg:p-5">
          <div className="grid gap-5 lg:min-h-[420px] lg:grid-cols-[0.8fr_1.2fr]">
            <section className="flex h-full flex-col justify-between rounded-[1.5rem] px-4 py-4 sm:px-6 sm:py-5 lg:px-7 lg:py-7">
              <div className="space-y-10 lg:space-y-14">
                <div className="flex items-center gap-5">
                  <div className="inline-flex h-14 w-14 shrink-0 items-center justify-center overflow-hidden rounded-[1.25rem] border border-white/12 bg-white/8 shadow-[0_0_30px_rgba(37,99,235,0.16)] backdrop-blur-xl sm:h-16 sm:w-16">
                    <img
                      src="/workspace/assets/logo_silhouette.png"
                      alt="Earthsoft logo"
                      className="h-full w-full object-cover"
                    />
                  </div>
                  <div className="min-w-0">
                    <div className="mb-2 text-[10px] font-semibold uppercase tracking-[0.36em] text-sky-300/90 sm:text-xs">
                      Earthwork Suite
                    </div>
                    <div className="flex flex-col">
                      <span className="text-3xl font-bold leading-none tracking-[-0.07em] text-white sm:text-4xl">
                        Earthsoft
                      </span>
                      <span className="mt-1 text-[10px] font-medium tracking-[0.02em] text-white/68 sm:text-xs">
                        The Railway earthwork workspace.
                      </span>
                    </div>
                  </div>
                </div>

                <p className="max-w-2xl text-sm leading-[1.7] text-white/72 sm:text-base">
                  Sign in to open the project workspace and continue with earthwork design.
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

            <section className="rounded-[1.75rem] border border-white/8 bg-[rgba(1,8,24,0.34)] p-4 shadow-2xl shadow-black/20 backdrop-blur-lg sm:p-5 lg:p-5">
              <form className="space-y-6" onSubmit={handleSubmit}>
                <div className="space-y-2">
                  <div className="inline-flex items-center gap-2 rounded-full border border-white/10 bg-white/5 px-3 py-1 text-xs uppercase tracking-[0.24em] text-white/55">
                    <LockKeyhole className="h-3.5 w-3.5" />
                    Secure Login
                  </div>
                  <h2 className="text-2xl font-semibold tracking-[-0.04em]">Workspace access</h2>
                  <p className="text-sm leading-6 text-white/60">
                    Use your authorized project login to open the Earthsoft dashboard.
                  </p>
                </div>

                <div className="space-y-4">
                  <label className="block space-y-2">
                    <span className="text-sm font-medium text-white/70">Username</span>
                    <input
                      className="w-full rounded-2xl border border-white/8 bg-[rgba(5,17,48,0.32)] px-4 py-2.5 text-white outline-none transition focus:border-blue-400/70 focus:bg-[rgba(7,23,63,0.52)] focus:ring-4 focus:ring-blue-500/15"
                      value={username}
                      onChange={(event) => setUsername(event.target.value)}
                      autoComplete="username"
                      placeholder="Enter your username"
                    />
                  </label>

                  <label className="block space-y-2">
                    <span className="text-sm font-medium text-white/70">Password</span>
                    <input
                      className="w-full rounded-2xl border border-white/8 bg-[rgba(5,17,48,0.32)] px-4 py-2.5 text-white outline-none transition focus:border-blue-400/70 focus:bg-[rgba(7,23,63,0.52)] focus:ring-4 focus:ring-blue-500/15"
                      value={password}
                      onChange={(event) => setPassword(event.target.value)}
                      autoComplete="current-password"
                      placeholder="Enter your password"
                      type="password"
                    />
                  </label>
                </div>

                <div className="min-h-6 text-sm text-red-300">{error}</div>

                <div className="rounded-[1.5rem] border border-blue-500/20 bg-[rgba(7,23,63,0.14)] p-4 text-white/80">
                  <div className="text-lg font-semibold text-white">Demo Login</div>
                  <div className="mt-3 space-y-2 text-sm">
                    <p>
                      Username: <span className="font-semibold text-white">{AUTH_USERNAME}</span>
                    </p>
                    <p>
                      Password: <span className="font-semibold text-white">{AUTH_PASSWORD}</span>
                    </p>
                  </div>
                </div>

                <button
                  type="submit"
                  disabled={isSubmitting}
                  className="inline-flex w-full items-center justify-center gap-2 rounded-2xl bg-[#3147ff]/92 px-4 py-2.5 font-semibold text-white transition hover:bg-[#4156ff]/92 disabled:cursor-not-allowed disabled:opacity-70"
                >
                  <LogIn className="h-4 w-4" />
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
