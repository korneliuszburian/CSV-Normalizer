module.exports = {
  mode: 'jit', // Just-In-Time mode
  content: [
    "./templates/**/*.html", // Adjust paths as needed
    "./static/src/**/*.js",
    "./node_modules/flowbite/**/*.js"
  ],
  theme: {
    extend: {},
  },
  plugins: [
    require("flowbite/plugin")
  ],
}
