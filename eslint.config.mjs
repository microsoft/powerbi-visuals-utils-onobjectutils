import js from "@eslint/js";
import tseslint from "typescript-eslint";
import powerbiVisualsPlugin from "eslint-plugin-powerbi-visuals";

export default [
    {
        ignores: [
            "node_modules/**",
            "dist/**",
            "coverage/**",
            "lib/**",
            ".tmp/**",
            "eslint.config.mjs"
        ]
    },
    js.configs.recommended,
    ...tseslint.configs.recommended,
    powerbiVisualsPlugin.configs.recommended,
    {
        files: ["**/*.ts", "**/*.mts"],
        rules: {
            "@typescript-eslint/explicit-module-boundary-types": "off",
            "@typescript-eslint/no-explicit-any": "off",
            "@typescript-eslint/no-inferrable-types": "off"
        }
    }
];
