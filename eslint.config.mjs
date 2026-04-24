import js from '@eslint/js';
import gas from 'eslint-plugin-googleappsscript';

export default [
  js.configs.recommended,
  {
    files: ['**/*.gs', '**/*.js'],
    plugins: {
      googleappsscript: gas,
    },
    languageOptions: {
      ecmaVersion: 2020,
      globals: {
        ...gas.environments.googleappsscript.globals,
      },
    },
    rules: {
      'no-unused-vars': ['error', { varsIgnorePattern: '^[A-Z_]' }],
      'no-undef': 'error',
      eqeqeq: 'error',
      'no-var': 'error',
      'prefer-const': 'error',
    },
  },
];
