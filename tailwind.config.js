/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./src/**/*.{js,jsx,ts,tsx}",
    "./src/taskpane/**/*.{js,jsx,ts,tsx}",
    "./src/client/**/*.{js,jsx,ts,tsx}",
  ],
  darkMode: 'class',
  theme: {
    extend: {
      colors: {
        border: "rgba(255, 255, 255, 0.1)",
        input: "rgba(255, 255, 255, 0.1)",
        ring: "rgba(76, 194, 255, 0.3)",
        background: "#000000",
        foreground: "#FFFFFF",
        primary: {
          DEFAULT: "rgba(76, 194, 255, 1)",
          foreground: "#000000",
        },
        secondary: {
          DEFAULT: "rgba(40, 40, 40, 0.7)",
          foreground: "#FFFFFF",
        },
        muted: {
          DEFAULT: "rgba(40, 40, 40, 0.7)",
          foreground: "rgba(255, 255, 255, 0.7)",
        },
      },
      borderRadius: {
        lg: "0.5rem",
        md: "calc(0.5rem - 2px)",
        sm: "calc(0.5rem - 4px)",
      },
      fontFamily: {
        roboto: ['Roboto', 'sans-serif'],
      },
    },
  },
  plugins: [],
}
