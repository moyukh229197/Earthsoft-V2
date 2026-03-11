import path from "node:path"
import { fileURLToPath } from "node:url"
import { defineConfig } from "vite"
import react from "@vitejs/plugin-react"
import tailwindcss from "@tailwindcss/vite"

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const githubRepo = process.env.GITHUB_REPOSITORY?.split("/")[1]
const base =
  process.env.GITHUB_ACTIONS && githubRepo ? `/${githubRepo}/` : "/"

export default defineConfig({
  base,
  plugins: [react(), tailwindcss()],
  publicDir: "public",
  server: {
    port: 8080,
  },
  build: {
    sourcemap: false,
  },
  resolve: {
    alias: {
      "@": path.resolve(__dirname),
    },
  },
})
