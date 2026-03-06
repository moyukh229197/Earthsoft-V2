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
        <div className="w-full max-w-7xl rounded-[2rem] border border-white/10 bg-[linear-gradient(135deg,rgba(8,24,66,0.88),rgba(2,10,28,0.78)_40%,rgba(2,8,20,0.84)_100%)] p-4 shadow-2xl shadow-black/60 backdrop-blur-2xl sm:p-6 lg:p-8">
          <div className="grid gap-6 lg:min-h-[760px] lg:grid-cols-[1.24fr_0.76fr]">
            <section className="flex h-full flex-col justify-between rounded-[1.75rem] px-5 py-5 sm:px-8 sm:py-7 lg:px-10 lg:py-10">
              <div className="space-y-10 lg:space-y-14">
                <div className="flex items-center gap-5">
                  <div className="inline-flex h-24 w-24 shrink-0 items-center justify-center overflow-hidden rounded-[2rem] border border-white/12 bg-white/8 shadow-[0_0_40px_rgba(37,99,235,0.2)] backdrop-blur-xl">
                    <img
                      src="/workspace/assets/logo_silhouette.png"
                      alt="Earthsoft logo"
                      className="h-full w-full object-cover"
                    />
                  </div>
                  <div className="min-w-0">
                    <div className="mb-3 text-sm font-semibold uppercase tracking-[0.42em] text-sky-300/90">
                      Earthwork Suite
                    </div>
                    <div className="flex flex-col">
                      <span className="text-5xl font-bold leading-none tracking-[-0.07em] text-white sm:text-6xl lg:text-7xl">
                        Earthsoft
                      </span>
                      <span className="mt-2 text-sm font-medium tracking-[0.02em] text-white/68 sm:text-base">
                        The Railway earthwork workspace.
                      </span>
                    </div>
                  </div>
                </div>

                <p className="max-w-3xl text-xl leading-[1.8] text-white/72 sm:text-2xl">
                  Sign in to open the project workspace and continue with earthwork design, verification, and estimates.
                </p>
              </div>

              <div className="flex flex-wrap gap-4">
                <div className="inline-flex items-center gap-3 rounded-full border border-white/10 bg-white/3 px-5 py-4 text-white/78 backdrop-blur-md">
                  <ShieldCheck className="h-5 w-5 text-white/70" />
                  <span className="text-sm sm:text-base">Session-based local access</span>
                </div>
                <div className="inline-flex items-center gap-3 rounded-full border border-white/10 bg-white/3 px-5 py-4 text-white/78 backdrop-blur-md">
                  <ChartNoAxesColumn className="h-5 w-5 text-white/70" />
                  <span className="text-sm sm:text-base">Railway earthwork workspace</span>
                </div>
              </div>
            </section>

            <section className="rounded-[2rem] border border-white/10 bg-[rgba(1,8,24,0.58)] p-5 shadow-2xl shadow-black/30 backdrop-blur-xl sm:p-6 lg:mx-4 lg:p-7">
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
                      className="w-full rounded-2xl border border-white/10 bg-[rgba(5,17,48,0.5)] px-4 py-3 text-white outline-none transition focus:border-blue-400/80 focus:bg-[rgba(7,23,63,0.72)] focus:ring-4 focus:ring-blue-500/20"
                      value={username}
                      onChange={(event) => setUsername(event.target.value)}
                      autoComplete="username"
                      placeholder="Enter your username"
                    />
                  </label>

                  <label className="block space-y-2">
                    <span className="text-sm font-medium text-white/70">Password</span>
                    <input
                      className="w-full rounded-2xl border border-white/10 bg-[rgba(5,17,48,0.5)] px-4 py-3 text-white outline-none transition focus:border-blue-400/80 focus:bg-[rgba(7,23,63,0.72)] focus:ring-4 focus:ring-blue-500/20"
                      value={password}
                      onChange={(event) => setPassword(event.target.value)}
                      autoComplete="current-password"
                      placeholder="Enter your password"
                      type="password"
                    />
                  </label>
                </div>

                <div className="min-h-6 text-sm text-red-300">{error}</div>

                <div className="rounded-[1.75rem] border border-blue-500/25 bg-[rgba(7,23,63,0.24)] p-5 text-white/80">
                  <div className="text-xl font-semibold text-white">Demo Login</div>
                  <div className="mt-3 space-y-2 text-base">
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
                  className="inline-flex w-full items-center justify-center gap-2 rounded-2xl bg-[#3147ff] px-4 py-3 font-semibold text-white transition hover:bg-[#4156ff] disabled:cursor-not-allowed disabled:opacity-70"
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
