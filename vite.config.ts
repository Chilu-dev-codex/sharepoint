import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    allowedHosts: ['equal-agrees-wiring-stopped.trycloudflare.com'],
  },
})
