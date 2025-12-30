import eslint from '@eslint/js';
import stylistic from '@stylistic/eslint-plugin';
import importPlugin from 'eslint-plugin-import';
import unusedImports from 'eslint-plugin-unused-imports';
import tseslint from 'typescript-eslint';

export default tseslint.config(
  eslint.configs.recommended,
  ...tseslint.configs.recommended,
  {
    plugins: {
      '@stylistic': stylistic,
      import: importPlugin,
      'unused-imports': unusedImports,
    },
    languageOptions: {
      parserOptions: {
        project: './tsconfig.json',
      },
    },
    settings: {
      'import/resolver': {
        typescript: {
          alwaysTryTypes: true,
          project: './tsconfig.json',
        },
      },
    },
    rules: {
      // Stylistic rules
      '@stylistic/quotes': ['error', 'single', { avoidEscape: true }],
      '@stylistic/semi': ['error', 'always'],
      '@stylistic/comma-dangle': ['error', 'always-multiline'],
      '@stylistic/indent': ['error', 2],
      '@stylistic/no-trailing-spaces': 'error',
      '@stylistic/eol-last': ['error', 'always'],
      '@stylistic/no-multiple-empty-lines': ['error', { max: 1, maxEOF: 0 }],
      '@stylistic/object-curly-spacing': ['error', 'always'],

      // Prefer path aliases over relative imports
      'no-restricted-imports': [
        'error',
        {
          patterns: [
            {
              group: ['../xlsx/*', '../sheet/*', '../utils/*', '../zip/*', '../xml/*'],
              message: 'Use path aliases (@xlsx/*, @sheet/*, @utils/*, @zip/*, @xml/*) instead of relative imports',
            },
            {
              group: ['../../xlsx/*', '../../sheet/*', '../../utils/*', '../../zip/*', '../../xml/*'],
              message: 'Use path aliases (@xlsx/*, @sheet/*, @utils/*, @zip/*, @xml/*) instead of relative imports',
            },
            {
              group: ['../../../xlsx/*', '../../../sheet/*', '../../../utils/*', '../../../zip/*', '../../../xml/*'],
              message: 'Use path aliases (@xlsx/*, @sheet/*, @utils/*, @zip/*, @xml/*) instead of relative imports',
            },
          ],
        },
      ],

      // Import order and organization
      'import/order': [
        'error',
        {
          groups: [
            'builtin',
            'external',
            'internal',
            ['parent', 'sibling', 'index'],
          ],
          pathGroups: [
            {
              pattern: '@xlsx/**',
              group: 'internal',
              position: 'before',
            },
            {
              pattern: '@sheet/**',
              group: 'internal',
              position: 'before',
            },
            {
              pattern: '@utils/**',
              group: 'internal',
              position: 'before',
            },
            {
              pattern: '@zip/**',
              group: 'internal',
              position: 'before',
            },
            {
              pattern: '@xml/**',
              group: 'internal',
              position: 'before',
            },
          ],
          pathGroupsExcludedImportTypes: ['builtin', 'external'],
          'newlines-between': 'never',
          alphabetize: {
            order: 'asc',
            caseInsensitive: true,
          },
        },
      ],
      // Require one newline after imports
      'import/newline-after-import': 'error',

      // Remove unused imports on --fix
      'no-unused-vars': 'off', // Turn off base rule
      '@typescript-eslint/no-unused-vars': 'off', // Turn off TypeScript rule
      'unused-imports/no-unused-imports': 'error', // Error on unused imports
      'unused-imports/no-unused-vars': [
        'warn',
        {
          vars: 'all',
          varsIgnorePattern: '^_',
          args: 'after-used',
          argsIgnorePattern: '^_',
        },
      ],
    },
  },
);
