import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// 把 'listing-expander' 換成你的 GitHub 倉庫名稱
export default defineConfig({
  plugins: [react()],
  base: '/listing-expander/',
})
