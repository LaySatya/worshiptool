/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,jsx,ts,tsx}'],
  theme: {
    extend: {
      fontFamily: {
        sans: ['Inter', 'system-ui', 'sans-serif'],
      },
      colors: {
        brand: {
          50:  '#f0efff',
          100: '#e3e1ff',
          200: '#cbc8ff',
          300: '#a9a4ff',
          400: '#8278ff',
          500: '#5b50e8',
          600: '#4a40d4',
          700: '#3b32b0',
          800: '#302890',
          900: '#2a2475',
        },
      },
      boxShadow: {
        'card': '0 1px 4px rgba(0,0,0,.06), 0 1px 2px rgba(0,0,0,.04)',
        'panel': '0 4px 24px rgba(0,0,0,.08)',
      },
    },
  },
  plugins: [],
}
