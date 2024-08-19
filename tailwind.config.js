/** @type {import('tailwindcss').Config} */
module.exports = {
  plugins: [
    require('flowbite/plugin')],

  content: ["./templates/*.html", "./node_modules/flowbite/**/*.js"],
  theme: {
    extend: {},
  },
  plugins: [],
}

