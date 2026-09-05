// @ts-check

import eslint from '@eslint/js';
import tseslint from 'typescript-eslint';

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
      '@typescript-eslint/naming-convention': [
        'error',
        {
          selector: 'objectLiteralProperty',
          format: ['camelCase'],
          filter: {
            regex: '^(Accept|Connection|Cookie|User-Agent|Content-Type|Accept-Language|__RequestVerificationToken|Authorization|Symbol|ClosePrice)$',
            match: false
          }
        }
      ]
    }
  }
);
