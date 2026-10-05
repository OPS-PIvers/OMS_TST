/**
 * Tailwind configuration for Index.html.
 *
 * The page used to load Tailwind's Play CDN, which downloads ~400 KB of
 * JavaScript on every visit and rebuilds the stylesheet in the browser each time
 * the page changes. The CSS is now built ahead of time by scripts/build-css.js
 * and inlined into Index.html. This theme is exactly the one the CDN was given.
 *
 * After changing classes in Index.html, run `npm run build:css` (CI fails if the
 * inlined CSS is out of date).
 */
module.exports = {
  // scripts/build-css.js passes Index.html in as raw content (minus the generated
  // block, so the CSS never feeds back into its own build).
  content: [],
  theme: {
    extend: {
      fontFamily: {
        sans: ['Lexend', 'sans-serif'],
      },
      colors: {
        ops: {
          blue: '#2d3f89',       // Primary Blue
          'blue-dark': '#1d2a5d', // Darkest Blue
          'blue-light': '#4356a0',
          'blue-lighter': '#eaecf5', // Info Background
          red: '#ad2122',        // Primary Red
          'red-dark': '#7a1718',
          'red-light': '#c13435',
          'red-lighter': '#e5c7c7', // Warning Background
          yellow: {
            dark: '#854d0e',
            lighter: '#fef9c3'
          },
          gray: {
            darkest: '#1a1a1a',
            dark: '#333333',
            DEFAULT: '#666666',
            light: '#999999',
            lighter: '#CCCCCC',
            lightest: '#f2f2f3'
          }
        }
      }
    }
  }
};
