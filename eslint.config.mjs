import eslint from '@eslint/js';
import tseslint from 'typescript-eslint';
import prettier from 'eslint-config-prettier';
import globals from 'globals';
import unusedImports from 'eslint-plugin-unused-imports';

export default tseslint.config(
  {
    ignores: [
      'dist/',
      'out/',
      'docs/',
      'website/',
      'node_modules/',
      '__customer__/',
      'src/dev.ts',
      'src/dev-*.ts',
    ],
  },
  eslint.configs.recommended,
  ...tseslint.configs.recommended,
  prettier,
  {
    languageOptions: {
      globals: {
        ...globals.node,
        ...globals.jest,
      },
    },
    plugins: {
      'unused-imports': unusedImports,
    },
    rules: {
      // The codebase still has legacy `any`s; Phase 4 of the roadmap
      // tightens types. Keep visible as warnings until then.
      '@typescript-eslint/no-explicit-any': 'warn',
      // Auto-fixable removal of unused imports; unused locals/args stay
      // errors via no-unused-vars below.
      'unused-imports/no-unused-imports': 'error',
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
          caughtErrorsIgnorePattern: '^_',
        },
      ],
    },
  },
  {
    // Tests: unused `result` vars usually flag missing assertions (see
    // ROADMAP Phase 5). Keep them visible as warnings, not errors.
    files: ['__tests__/**/*.ts'],
    rules: {
      '@typescript-eslint/no-unused-vars': [
        'warn',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
          caughtErrorsIgnorePattern: '^_',
        },
      ],
    },
  },
);
