// eslint.config.js
import antfu from '@antfu/eslint-config'

export default antfu({
  // Type of the project. 'lib' for libraries, the default is 'app'
  type: 'app',

  // Enable stylistic formatting rules
  // stylistic: true,

  rules: {
    'no-console': 'off',
    'ts/no-namespace': 'off',
    'unicorn/prefer-number-properties': 'off',
    'node/prefer-global/buffer': 'off',
    'vue/custom-event-name-casing': 'off',
  },

  // Or customize the stylistic rules
  stylistic: {
    indent: 2, // 4, or 'tab'
    quotes: 'single', // or 'double'
  },

  // TypeScript and Vue are autodetected, you can also explicitly enable them:
  typescript: true,
  vue: true,

  // Disable jsonc and yaml support
  jsonc: false,
  yaml: false,

  // `.eslintignore` is no longer supported in Flat config, use `ignores` instead
  ignores: [
    '**/fixtures',
    // ...globs
  ],
})
