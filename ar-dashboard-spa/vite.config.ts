import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  // Use relative paths so the built site can be opened from file://
  base: './',
  plugins: [react()],
})
