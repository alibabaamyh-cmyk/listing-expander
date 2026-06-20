import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { resolve } from 'path'

export default defineConfig({
  plugins: [react()],
  base: '/listing-expander/',
  build: {
    rollupOptions: {
      input: {
        main: resolve(__dirname, 'index.html'),
        advisor: resolve(__dirname, 'advisor/index.html'),
      }
    }
  }
})
