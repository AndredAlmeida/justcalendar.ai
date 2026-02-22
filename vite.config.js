import { existsSync, readFileSync } from 'node:fs'
import { defineConfig } from 'vite'

const tlsKeyPath = new URL('./certs/justcal.ai.key', import.meta.url)
const tlsCertPath = new URL('./certs/justcal.ai.crt', import.meta.url)

const httpsConfig =
  existsSync(tlsKeyPath) && existsSync(tlsCertPath)
    ? {
        key: readFileSync(tlsKeyPath),
        cert: readFileSync(tlsCertPath),
      }
    : undefined

function createMalformedUriGuardPlugin() {
  const sanitizeUrl = (rawUrl) => {
    if (typeof rawUrl !== 'string' || rawUrl.length === 0) {
      return rawUrl
    }
    // Guard against malformed percent sequences that crash decodeURI in
    // Vite static middlewares.
    return rawUrl.replace(/%(?![0-9A-Fa-f]{2})/g, '%25')
  }

  const attachSanitizer = (middlewares) => {
    middlewares.use((req, _res, next) => {
      req.url = sanitizeUrl(req.url)
      next()
    })
  }

  return {
    name: 'malformed-uri-guard',
    configureServer(server) {
      attachSanitizer(server.middlewares)
    },
    configurePreviewServer(server) {
      attachSanitizer(server.middlewares)
    },
  }
}

export default defineConfig({
  plugins: [createMalformedUriGuardPlugin()],
  server: {
    host: '0.0.0.0',
    port: 443,
    strictPort: true,
    https: httpsConfig,
    allowedHosts: [
      'mbp',
      'mbp.tail1592c.ts.net',
      'justcal.ai',
      'www.justcal.ai',
      'justcalendar.ai',
      'www.justcalendar.ai',
    ],
  },
  preview: {
    host: '0.0.0.0',
    port: 443,
    strictPort: true,
    https: httpsConfig,
    allowedHosts: [
      'mbp',
      'mbp.tail1592c.ts.net',
      'justcal.ai',
      'www.justcal.ai',
      'justcalendar.ai',
      'www.justcalendar.ai',
    ],
  },
})
