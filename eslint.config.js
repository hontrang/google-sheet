// @ts-check

import eslint from '@eslint/js';
import tseslint from 'typescript-eslint';

const objectLiteralPropertyNaming = {
  selector: 'objectLiteralProperty',
  format: ['camelCase'],
  filter: {
    regex: '^(Accept|Connection|Cookie|User-Agent|Content-Type|Accept-Language|__RequestVerificationToken|Authorization|Symbol|ClosePrice)$',
    match: false
  }
};

export default tseslint.config(
  eslint.configs.recommended,
  ...tseslint.configs.recommended,
  ...tseslint.configs.strict,
  ...tseslint.configs.stylistic,
  ...tseslint.configs.recommendedTypeChecked,
  ...tseslint.configs.stylisticTypeChecked,
  {
    languageOptions: {
      parserOptions: {
        project: true,
        projectFolderIgnoreList: ['**/node_modules/**', '**/dist/**']
      }
    }
  },
  {
    files: ['**/*.ts'],
    ...tseslint.configs.disableTypeChecked
  },
  {
    languageOptions: {
      parserOptions: {
        project: true,
        tsconfigRootDir: import.meta.dirname
      }
    }
  },
  {
    rules: {
      '@typescript-eslint/array-type': 'off',
      '@typescript-eslint/no-namespace': 'off',
      '@typescript-eslint/naming-convention': ['error', objectLiteralPropertyNaming]
    }
  },
  {
    // Variable naming for the repo's own code (not the vendored src/assets/lib)
    files: ['src/**/*.ts', 'tests/**/*.ts', '*.ts'],
    ignores: ['src/assets/lib/**'],
    rules: {
      '@typescript-eslint/naming-convention': [
        'error',
        objectLiteralPropertyNaming,
        // Aliases of library globals keep the library's own casing
        {
          selector: 'variable',
          filter: { regex: '^QUnit$', match: true },
          format: null
        },
        // Local variables are camelCase even when `const`; UPPER_CASE is only for module-level constants
        {
          selector: 'variable',
          format: ['camelCase']
        },
        {
          selector: 'variable',
          modifiers: ['global', 'const'],
          format: ['camelCase', 'UPPER_CASE'],
          trailingUnderscore: 'allow'
        }
      ]
    }
  }
);
