import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'
import basicSsl from '@vitejs/plugin-basic-ssl'

// https://vitejs.dev/config/
export default defineConfig(({ mode }) => {
  // @ts-expect-error
  const env = loadEnv(mode, process.cwd());
  if (env.VITE_ENVIRONMENT === 'development') {
    return {
      plugins: [
        react(),
        basicSsl()
      ]
    }
  }


  return {
    plugins: [
      react(),
    ]
  }
})
