var _a;
import path from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
var __dirname = path.dirname(fileURLToPath(import.meta.url));
var githubRepo = (_a = process.env.GITHUB_REPOSITORY) === null || _a === void 0 ? void 0 : _a.split("/")[1];
var base = process.env.GITHUB_ACTIONS && githubRepo ? "/".concat(githubRepo, "/") : "/";
export default defineConfig({
    base: base,
    plugins: [react(), tailwindcss()],
    publicDir: "public",
    build: {
        sourcemap: false,
    },
    resolve: {
        alias: {
            "@": path.resolve(__dirname),
        },
    },
});
