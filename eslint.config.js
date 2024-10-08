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
                project: true
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
            '@typescript-eslint/naming-convention': 'error'
        }
    }
);